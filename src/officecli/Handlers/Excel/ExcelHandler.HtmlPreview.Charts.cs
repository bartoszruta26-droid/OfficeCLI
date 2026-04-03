// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Drawing = DocumentFormat.OpenXml.Drawing;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    /// <summary>
    /// Render all charts in a worksheet as SVG elements, respecting anchor positions.
    /// Charts with overlapping row ranges are placed side-by-side using flex layout.
    /// </summary>
    private void RenderSheetCharts(StringBuilder sb, WorksheetPart worksheetPart)
    {
        var drawingsPart = worksheetPart.DrawingsPart;
        if (drawingsPart?.WorksheetDrawing == null) return;

        // Find all graphic frames that contain chart references
        var chartFrames = drawingsPart.WorksheetDrawing
            .Descendants<XDR.GraphicFrame>()
            .Where(gf => gf.Descendants<C.ChartReference>().Any())
            .ToList();

        if (chartFrames.Count == 0) return;

        // Read anchor positions and group charts into rows (overlapping row ranges = same row)
        var chartAnchors = chartFrames.Select(gf =>
        {
            var anchor = gf.Parent as XDR.TwoCellAnchor;
            int fromRow = 0, toRow = 0, fromCol = 0;
            if (anchor?.FromMarker != null && anchor?.ToMarker != null)
            {
                int.TryParse(anchor.FromMarker.RowId?.Text, out fromRow);
                int.TryParse(anchor.ToMarker.RowId?.Text, out toRow);
                int.TryParse(anchor.FromMarker.ColumnId?.Text, out fromCol);
            }
            return (gf, fromRow, toRow, fromCol);
        }).OrderBy(x => x.fromRow).ThenBy(x => x.fromCol).ToList();

        // Group into rows: charts whose row ranges overlap go in the same flex row
        var rows = new List<List<XDR.GraphicFrame>>();
        int currentRowEnd = -1;
        List<XDR.GraphicFrame>? currentRow = null;
        foreach (var (gf, fromRow, toRow, _) in chartAnchors)
        {
            if (currentRow == null || fromRow >= currentRowEnd)
            {
                currentRow = new List<XDR.GraphicFrame>();
                rows.Add(currentRow);
                currentRowEnd = toRow;
            }
            else
            {
                currentRowEnd = Math.Max(currentRowEnd, toRow);
            }
            currentRow.Add(gf);
        }

        foreach (var row in rows)
        {
            if (row.Count > 1)
            {
                sb.AppendLine("<div style=\"display:flex;gap:16px;flex-wrap:wrap\">");
                foreach (var gf in row)
                    RenderExcelChart(sb, gf, drawingsPart, worksheetPart);
                sb.AppendLine("</div>");
            }
            else
            {
                RenderExcelChart(sb, row[0], drawingsPart, worksheetPart);
            }
        }
    }

    private void RenderExcelChart(StringBuilder sb, XDR.GraphicFrame gf,
        DrawingsPart drawingsPart, WorksheetPart worksheetPart)
    {
        // 1. Get chart reference and load ChartPart
        var chartRef = gf.Descendants<C.ChartReference>().FirstOrDefault();
        if (chartRef?.Id?.Value == null) return;

        C.Chart? chart;
        C.PlotArea? plotArea;
        try
        {
            var chartPart = (ChartPart)drawingsPart.GetPartById(chartRef.Id.Value);
            chart = chartPart.ChartSpace?.GetFirstChild<C.Chart>();
            plotArea = chart?.GetFirstChild<C.PlotArea>();
            if (plotArea == null) return;
        }
        catch { return; }

        // 2. Read chart data using shared ChartHelper
        var chartType = ChartHelper.DetectChartType(plotArea) ?? "bar";
        var categories = ChartHelper.ReadCategories(plotArea) ?? [];
        var seriesList = ChartHelper.ReadAllSeries(plotArea);
        if (seriesList.Count == 0) return;

        // 2b. Resolve series names from cell references when strCache is missing
        if (seriesList.Any(s => s.name == "?"))
        {
            var nameSerEls = plotArea.Descendants<OpenXmlCompositeElement>()
                .Where(e => e.LocalName == "ser" && e.Parent != null &&
                    (e.Parent.LocalName.Contains("Chart") || e.Parent.LocalName.Contains("chart")))
                .ToList();
            for (int i = 0; i < seriesList.Count && i < nameSerEls.Count; i++)
            {
                if (seriesList[i].name != "?") continue;
                var strRef = nameSerEls[i].GetFirstChild<C.SeriesText>()
                    ?.Descendants<C.StringReference>().FirstOrDefault();
                var formula = strRef?.GetFirstChild<C.Formula>()?.Text;
                if (!string.IsNullOrEmpty(formula))
                {
                    var resolved = ReadCellRangeAsStrings(formula);
                    if (resolved != null && resolved.Length > 0)
                        seriesList[i] = (resolved[0], seriesList[i].values);
                }
            }
        }

        // 2c. Resolve cell references when cache is missing (chart references other sheets)
        var needsCatResolve = categories.Length == 0;
        var needsValResolve = seriesList.All(s => s.values.Length == 0);
        if (needsCatResolve || needsValResolve)
        {
            ResolveChartDataFromCells(plotArea, ref categories, ref seriesList, needsCatResolve, needsValResolve);
            if (seriesList.All(s => s.values.Length == 0)) return;
        }

        // 3. Read series colors (and per-point colors for pie/doughnut)
        var seriesColors = new List<string>();
        var serElements = plotArea.Descendants<OpenXmlCompositeElement>()
            .Where(e => e.LocalName == "ser").ToList();
        var isPieType = chartType.Contains("pie") || chartType.Contains("doughnut");

        if (isPieType && serElements.Count > 0)
        {
            // Pie/doughnut: colors are per data point (dPt), not per series
            var ser = serElements[0];
            var dPts = ser.Elements<OpenXmlCompositeElement>().Where(e => e.LocalName == "dPt").ToList();
            var catCount = seriesList.FirstOrDefault().values?.Length ?? 0;
            for (int i = 0; i < catCount; i++)
            {
                var dPt = dPts.FirstOrDefault(d =>
                {
                    var idxEl = d.Elements<OpenXmlCompositeElement>().FirstOrDefault(e => e.LocalName == "idx");
                    if (idxEl == null) return false;
                    var valAttr = idxEl.GetAttributes().FirstOrDefault(a => a.LocalName == "val");
                    return valAttr.Value == i.ToString();
                });
                var spPr = dPt?.GetFirstChild<C.ChartShapeProperties>();
                var fill = spPr?.GetFirstChild<Drawing.SolidFill>();
                var rgb = fill?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
                seriesColors.Add(rgb != null ? $"#{rgb}" : ChartSvgRenderer.DefaultColors[i % ChartSvgRenderer.DefaultColors.Length]);
            }
        }
        else
        {
            // Determine which series belong to line/scatter charts (stroke color from a:ln)
            var lineSerIndices = new HashSet<int>();
            foreach (var chartEl in plotArea.Elements<OpenXmlCompositeElement>()
                .Where(e => e.LocalName is "lineChart" or "scatterChart"))
            {
                foreach (var s in chartEl.Elements<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser"))
                {
                    var idx = s.GetFirstChild<C.Index>()?.Val?.Value;
                    if (idx.HasValue) lineSerIndices.Add((int)idx.Value);
                }
            }

            for (int i = 0; i < seriesList.Count; i++)
            {
                var serEl = i < serElements.Count ? serElements[i] : null;
                var spPr = serEl?.GetFirstChild<C.ChartShapeProperties>();

                // For line/scatter series, prefer line stroke color (a:ln > a:solidFill)
                string? rgb = null;
                var serIdx = serEl?.GetFirstChild<C.Index>()?.Val?.Value;
                if (serIdx.HasValue && lineSerIndices.Contains((int)serIdx.Value))
                {
                    var ln = spPr?.GetFirstChild<Drawing.Outline>();
                    rgb = ln?.GetFirstChild<Drawing.SolidFill>()?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
                }
                // Fallback to solidFill (works for bar/area/pie)
                rgb ??= spPr?.GetFirstChild<Drawing.SolidFill>()?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
                seriesColors.Add(rgb != null ? $"#{rgb}" : ChartSvgRenderer.DefaultColors[i % ChartSvgRenderer.DefaultColors.Length]);
            }
        }

        // 4. Estimate chart dimensions from TwoCellAnchor
        var (widthPt, heightPt) = EstimateChartSize(gf);

        // 5. Read chart metadata
        var chartTitle = chart?.GetFirstChild<C.Title>();
        var titleText = chartTitle?.Descendants<Drawing.Text>().FirstOrDefault()?.Text ?? "";
        var titleFontSize = chartTitle?.Descendants<Drawing.RunProperties>().FirstOrDefault()?.FontSize;
        var titleSizeCss = titleFontSize?.HasValue == true ? $"{titleFontSize.Value / 100.0:0.##}pt" : "10pt";

        var dataLabels = plotArea.Descendants<C.DataLabels>().FirstOrDefault();
        var showValues = dataLabels?.GetFirstChild<C.ShowValue>()?.Val?.Value == true
            || dataLabels?.GetFirstChild<C.ShowCategoryName>()?.Val?.Value == true
            || dataLabels?.GetFirstChild<C.ShowPercent>()?.Val?.Value == true;

        var plotFillColor = plotArea.GetFirstChild<C.ShapeProperties>()
            ?.GetFirstChild<Drawing.SolidFill>()?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
        var chartFillColor = chart?.Parent?.GetFirstChild<C.ChartShapeProperties>()
            ?.GetFirstChild<Drawing.SolidFill>()?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;

        // Axis info
        var valAxis = plotArea.GetFirstChild<C.ValueAxis>();
        var valAxisTitle = valAxis?.GetFirstChild<C.Title>()?.Descendants<Drawing.Text>().FirstOrDefault()?.Text;
        var catAxis = plotArea.GetFirstChild<C.CategoryAxis>();
        var catAxisTitle = catAxis?.GetFirstChild<C.Title>()?.Descendants<Drawing.Text>().FirstOrDefault()?.Text;

        var valScaling = valAxis?.GetFirstChild<C.Scaling>();
        double? ooxmlAxisMax = null, ooxmlAxisMin = null, ooxmlMajorUnit = null;
        if (valScaling?.GetFirstChild<C.MaxAxisValue>()?.Val?.HasValue == true)
            ooxmlAxisMax = valScaling.GetFirstChild<C.MaxAxisValue>()!.Val!.Value;
        if (valScaling?.GetFirstChild<C.MinAxisValue>()?.Val?.HasValue == true)
            ooxmlAxisMin = valScaling.GetFirstChild<C.MinAxisValue>()!.Val!.Value;
        if (valAxis?.GetFirstChild<C.MajorUnit>()?.Val?.HasValue == true)
            ooxmlMajorUnit = valAxis.GetFirstChild<C.MajorUnit>()!.Val!.Value;

        var gapWidthEl = plotArea.Descendants<C.GapWidth>().FirstOrDefault();
        int? ooxmlGapWidth = gapWidthEl?.Val?.HasValue == true ? (int)gapWidthEl.Val.Value : null;

        var valAxisFontSize = valAxis?.Descendants<Drawing.RunProperties>().FirstOrDefault()?.FontSize;
        var catAxisFontSize = catAxis?.Descendants<Drawing.RunProperties>().FirstOrDefault()?.FontSize;
        int valLabelPx = valAxisFontSize?.HasValue == true ? (int)(valAxisFontSize.Value / 100.0 * 96 / 72) : 9;
        int catLabelPx = catAxisFontSize?.HasValue == true ? (int)(catAxisFontSize.Value / 100.0 * 96 / 72) : 9;

        // Legend
        var legendEl = chart?.GetFirstChild<C.Legend>();
        var isPieOrDoughnut = chartType.Contains("pie") || chartType.Contains("doughnut");
        bool hasLegend;
        if (legendEl != null)
        {
            var deleteEl = legendEl.GetFirstChild<C.Delete>();
            hasLegend = deleteEl?.Val?.Value != true;
        }
        else hasLegend = seriesList.Count > 1 || isPieOrDoughnut;

        // 6. Create renderer with Excel-appropriate colors (light background)
        var renderer = new ChartSvgRenderer
        {
            ValueColor = "#333",
            CatColor = "#555",
            AxisColor = "#666",
            GridColor = "#ddd",
            AxisLineColor = "#999",
            ValFontPx = valLabelPx,
            CatFontPx = catLabelPx
        };

        // 7. Build SVG
        var svgW = Math.Max(widthPt, 225);
        var svgH = Math.Max(heightPt, 150);
        var titleH = string.IsNullOrEmpty(titleText) ? 0 : 30;
        var legendH = hasLegend ? 30 : 0;
        var chartSvgH = svgH - titleH - legendH;

        int marginTop = 10, marginRight = 15, marginBottom = 30, marginLeft = 45;
        var plotW = svgW - marginLeft - marginRight;
        var plotH = chartSvgH - marginTop - marginBottom;
        if (plotW < 50 || plotH < 50) return;

        var bgStyle = chartFillColor != null ? $"background:#{chartFillColor};" : "";
        sb.AppendLine($"<div class=\"chart-container\" style=\"max-width:{svgW}pt;flex:1;min-width:200pt;{bgStyle}\">");

        // Title
        if (!string.IsNullOrEmpty(titleText))
            sb.AppendLine($"  <div style=\"text-align:center;font-size:{titleSizeCss};font-weight:bold;padding:6px 0;color:#333\">{HtmlEncode(titleText)}</div>");

        sb.AppendLine($"  <svg viewBox=\"0 0 {svgW} {chartSvgH}\" style=\"width:100%;height:auto\" preserveAspectRatio=\"xMidYMin meet\">");

        // Plot area background
        if (plotFillColor != null)
            sb.AppendLine($"    <rect x=\"{marginLeft}\" y=\"{marginTop}\" width=\"{plotW}\" height=\"{plotH}\" fill=\"#{plotFillColor}\"/>");

        // Route to correct chart renderer
        var is3D = chartType.Contains("3d");
        if (chartType.Contains("pie") || chartType.Contains("doughnut"))
        {
            var isDoughnut = chartType.Contains("doughnut");
            var holeSize = 0.0;
            if (isDoughnut)
            {
                var holeSizeEl = plotArea.Descendants<C.HoleSize>().FirstOrDefault();
                holeSize = (holeSizeEl?.Val?.Value ?? 50) / 100.0;
            }
            renderer.RenderPieChartSvg(sb, seriesList, categories, seriesColors, svgW, chartSvgH, holeSize, showValues);
        }
        else if (chartType.Contains("area"))
        {
            var areaStacked = chartType.Contains("stacked") || chartType.Contains("Stacked");
            renderer.RenderAreaChartSvg(sb, seriesList, categories, seriesColors, marginLeft, marginTop, plotW, plotH, areaStacked);
        }
        else if (chartType == "combo")
        {
            renderer.RenderComboChartSvg(sb, plotArea, seriesList, categories, seriesColors, marginLeft, marginTop, plotW, plotH);
        }
        else if (chartType.Contains("radar"))
        {
            renderer.RenderRadarChartSvg(sb, seriesList, categories, seriesColors, svgW, chartSvgH, catLabelPx);
        }
        else if (chartType == "bubble")
        {
            renderer.RenderBubbleChartSvg(sb, plotArea, seriesList, categories, seriesColors, marginLeft, marginTop, plotW, plotH);
        }
        else if (chartType == "stock")
        {
            renderer.RenderStockChartSvg(sb, plotArea, seriesList, categories, seriesColors, marginLeft, marginTop, plotW, plotH);
        }
        else if (chartType.Contains("line") || chartType == "scatter")
        {
            renderer.RenderLineChartSvg(sb, seriesList, categories, seriesColors, marginLeft, marginTop, plotW, plotH, showValues);
        }
        else
        {
            // Column/bar variants
            var isHorizontal = chartType.Contains("bar") && !chartType.Contains("column");
            var isStacked = chartType.Contains("stacked") || chartType.Contains("Stacked");
            var isPercent = chartType.Contains("percent") || chartType.Contains("Percent");
            renderer.RenderBarChartSvg(sb, seriesList, categories, seriesColors, marginLeft, marginTop, plotW, plotH,
                isHorizontal, isStacked, isPercent, ooxmlAxisMax, ooxmlAxisMin, ooxmlMajorUnit, ooxmlGapWidth, valLabelPx, catLabelPx, showValues);
        }

        // Axis titles inside SVG
        if (!string.IsNullOrEmpty(valAxisTitle))
            sb.AppendLine($"    <text x=\"10\" y=\"{chartSvgH / 2}\" fill=\"#666\" font-size=\"{valLabelPx}\" text-anchor=\"middle\" dominant-baseline=\"middle\" transform=\"rotate(-90,10,{chartSvgH / 2})\">{HtmlEncode(valAxisTitle)}</text>");
        if (!string.IsNullOrEmpty(catAxisTitle))
            sb.AppendLine($"    <text x=\"{svgW / 2}\" y=\"{chartSvgH - 2}\" fill=\"#666\" font-size=\"{valLabelPx}\" text-anchor=\"middle\">{HtmlEncode(catAxisTitle)}</text>");

        sb.AppendLine("  </svg>");

        // Legend
        if (hasLegend)
        {
            var legendFontSize = legendEl?.Descendants<Drawing.RunProperties>().FirstOrDefault()?.FontSize;
            var legendSizeCss = legendFontSize?.HasValue == true ? $"{legendFontSize.Value / 100.0:0.##}pt" : "8pt";
            sb.Append($"  <div style=\"display:flex;justify-content:center;gap:16px;padding:6px 0;font-size:{legendSizeCss};color:#555\">");
            if (isPieOrDoughnut && categories.Length > 0)
            {
                for (int i = 0; i < categories.Length; i++)
                {
                    var color = i < seriesColors.Count ? seriesColors[i] : ChartSvgRenderer.DefaultColors[i % ChartSvgRenderer.DefaultColors.Length];
                    sb.Append($"<span style=\"display:inline-flex;align-items:center;gap:4px\"><span style=\"display:inline-block;width:12px;height:12px;background:{color};border-radius:1px\"></span>{HtmlEncode(categories[i])}</span>");
                }
            }
            else
            {
                for (int i = 0; i < seriesList.Count && i < seriesColors.Count; i++)
                    sb.Append($"<span style=\"display:inline-flex;align-items:center;gap:4px\"><span style=\"display:inline-block;width:12px;height:12px;background:{seriesColors[i]};border-radius:1px\"></span>{HtmlEncode(seriesList[i].name)}</span>");
            }
            sb.AppendLine("</div>");
        }

        sb.AppendLine("</div>");
    }

    /// <summary>
    /// Estimate chart pixel size from the TwoCellAnchor parent.
    /// </summary>
    private static (int widthPt, int heightPt) EstimateChartSize(XDR.GraphicFrame gf)
    {
        var anchor = gf.Parent as XDR.TwoCellAnchor;
        if (anchor == null) return (450, 263);

        var from = anchor.FromMarker;
        var to = anchor.ToMarker;
        if (from == null || to == null) return (450, 263);

        var fromCol = int.TryParse(from.ColumnId?.Text, out var fc) ? fc : 0;
        var toCol = int.TryParse(to.ColumnId?.Text, out var tc) ? tc : 0;
        var fromRow = int.TryParse(from.RowId?.Text, out var fr) ? fr : 0;
        var toRow = int.TryParse(to.RowId?.Text, out var tr) ? tr : 0;

        var fromColOff = long.TryParse(from.ColumnOffset?.Text, out var fco) ? fco : 0;
        var toColOff = long.TryParse(to.ColumnOffset?.Text, out var tco) ? tco : 0;
        var fromRowOff = long.TryParse(from.RowOffset?.Text, out var fro) ? fro : 0;
        var toRowOff = long.TryParse(to.RowOffset?.Text, out var tro) ? tro : 0;

        // Default column width ~48pt, default row height ~15pt; offsets in EMU (1pt = 12700 EMU)
        double totalWidth = (toCol - fromCol) * 48.0 + (toColOff - fromColOff) / 12700.0;
        double totalHeight = (toRow - fromRow) * 15.0 + (toRowOff - fromRowOff) / 12700.0;

        return ((int)Math.Max(totalWidth, 225), (int)Math.Max(totalHeight, 150));
    }

    /// <summary>
    /// Resolve chart data from actual cells when the chart XML has no cache.
    /// Parses formula references like "'Income Statement'!$B$10:$D$10" and reads cell values.
    /// </summary>
    private void ResolveChartDataFromCells(C.PlotArea plotArea,
        ref string[] categories, ref List<(string name, double[] values)> seriesList,
        bool resolveCats, bool resolveVals)
    {
        if (resolveCats)
        {
            var catRef = ChartHelper.ReadCategoriesRef(plotArea);
            if (catRef != null)
            {
                var resolved = ReadCellRangeAsStrings(catRef);
                if (resolved != null) categories = resolved;
            }
        }

        if (resolveVals)
        {
            var newSeries = new List<(string name, double[] values)>();
            foreach (var ser in plotArea.Descendants<OpenXmlCompositeElement>()
                .Where(e => e.LocalName == "ser" && e.Parent != null &&
                    (e.Parent.LocalName.Contains("Chart") || e.Parent.LocalName.Contains("chart"))))
            {
                var serText = ser.GetFirstChild<C.SeriesText>();
                var name = serText?.Descendants<C.NumericValue>().FirstOrDefault()?.Text ?? "?";

                var valRef = ChartHelper.ReadFormulaRef(ser.GetFirstChild<C.Values>())
                    ?? ChartHelper.ReadFormulaRef(ser.Elements<OpenXmlCompositeElement>()
                        .FirstOrDefault(e => e.LocalName == "yVal"));

                double[] values = [];
                if (valRef != null)
                {
                    var resolved = ReadCellRangeAsDoubles(valRef);
                    if (resolved != null) values = resolved;
                }

                newSeries.Add((name, values));
            }
            if (newSeries.Count > 0) seriesList = newSeries;
        }
    }

    /// <summary>
    /// Parse a cell range reference like "'Sheet Name'!$B$1:$D$1" and return cell values as strings.
    /// </summary>
    private string[]? ReadCellRangeAsStrings(string formula)
    {
        var (sheetData, startCol, startRow, endCol, endRow) = ParseCellRangeFormula(formula);
        if (sheetData == null) return null;

        var results = new List<string>();
        for (int r = startRow; r <= endRow; r++)
        {
            for (int c = startCol; c <= endCol; c++)
            {
                var cellRef = GetColumnLetter(c) + r;
                var cell = sheetData.Descendants<Cell>()
                    .FirstOrDefault(cl => cl.CellReference?.Value == cellRef);
                results.Add(cell != null ? GetCellDisplayValue(cell) : "");
            }
        }
        return results.ToArray();
    }

    /// <summary>
    /// Parse a cell range reference and return cell values as doubles.
    /// Uses FormulaEvaluator with cross-sheet support.
    /// </summary>
    private double[]? ReadCellRangeAsDoubles(string formula)
    {
        var (sheetData, startCol, startRow, endCol, endRow) = ParseCellRangeFormula(formula);
        if (sheetData == null) return null;

        var evaluator = new Core.FormulaEvaluator(sheetData, _doc.WorkbookPart);
        var results = new List<double>();
        for (int r = startRow; r <= endRow; r++)
        {
            for (int c = startCol; c <= endCol; c++)
            {
                var cellRef = GetColumnLetter(c) + r;
                var cell = sheetData.Descendants<Cell>()
                    .FirstOrDefault(cl => cl.CellReference?.Value == cellRef);

                double val = 0;
                if (cell != null)
                {
                    var raw = cell.CellValue?.Text;
                    if (!string.IsNullOrEmpty(raw) && double.TryParse(raw,
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out var v))
                    {
                        val = v;
                    }
                    else if (cell.CellFormula?.Text != null)
                    {
                        val = evaluator.TryEvaluate(cell.CellFormula.Text) ?? 0;
                    }
                }
                results.Add(val);
            }
        }
        return results.ToArray();
    }

    /// <summary>
    /// Parse "'Sheet Name'!$B$1:$D$1" into (SheetData, startCol, startRow, endCol, endRow).
    /// </summary>
    private (SheetData? sheetData, int startCol, int startRow, int endCol, int endRow) ParseCellRangeFormula(string formula)
    {
        // Pattern: optional 'SheetName'! or SheetName! prefix, then cell range like $B$1:$D$1 or B1:D1
        var match = Regex.Match(formula, @"^(?:'([^']+)'|([^!]+))!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$");
        if (!match.Success) return (null, 0, 0, 0, 0);

        var sheetName = match.Groups[1].Success ? match.Groups[1].Value : match.Groups[2].Value;
        var startColStr = match.Groups[3].Value;
        var startRow = int.Parse(match.Groups[4].Value);
        var endColStr = match.Groups[5].Success ? match.Groups[5].Value : startColStr;
        var endRow = match.Groups[6].Success ? int.Parse(match.Groups[6].Value) : startRow;

        var startCol = ColumnLetterToIndex(startColStr);
        var endCol = ColumnLetterToIndex(endColStr);

        // Find the worksheet by name
        var workbookPart = _doc.WorkbookPart;
        if (workbookPart == null) return (null, 0, 0, 0, 0);

        var sheet = workbookPart.Workbook?.Descendants<Sheet>()
            .FirstOrDefault(s => s.Name?.Value == sheetName);
        if (sheet?.Id?.Value == null) return (null, 0, 0, 0, 0);

        try
        {
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id.Value);
            var sheetData = worksheetPart.Worksheet?.GetFirstChild<SheetData>();
            return (sheetData, startCol, startRow, endCol, endRow);
        }
        catch { return (null, 0, 0, 0, 0); }
    }

    private static int ColumnLetterToIndex(string col)
    {
        int result = 0;
        foreach (var c in col)
            result = result * 26 + (c - 'A' + 1);
        return result;
    }

    private static string GetColumnLetter(int colIndex)
    {
        var result = "";
        while (colIndex > 0)
        {
            colIndex--;
            result = (char)('A' + colIndex % 26) + result;
            colIndex /= 26;
        }
        return result;
    }
}
