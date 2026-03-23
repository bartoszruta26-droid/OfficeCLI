// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Reflection;

namespace OfficeCli.Core;

/// <summary>
/// Installs officecli skills (SKILL.md) into AI client skill directories.
/// </summary>
public static class SkillInstaller
{
    public static void Install(string target)
    {
        switch (target.ToLowerInvariant())
        {
            case "claude" or "claude-code":
                InstallTo("Claude Code", Path.Combine(Home, ".claude", "skills", "officecli", "SKILL.md"));
                break;
            case "copilot" or "github-copilot":
                InstallTo("GitHub Copilot", Path.Combine(Home, ".copilot", "skills", "officecli", "SKILL.md"));
                break;
            case "codex" or "openai-codex":
                InstallTo("Codex CLI", Path.Combine(Home, ".agents", "skills", "officecli", "SKILL.md"));
                break;
            case "all":
                Install("claude");
                Install("copilot");
                Install("codex");
                break;
            default:
                Console.Error.WriteLine($"Unknown target: {target}");
                Console.Error.WriteLine("Supported: claude, copilot, codex, all");
                break;
        }
    }

    private static void InstallTo(string displayName, string targetPath)
    {
        var content = LoadEmbeddedResource("OfficeCli.Resources.skill-officecli.md");
        if (content == null)
        {
            Console.Error.WriteLine($"  {displayName}: embedded resource not found");
            return;
        }

        if (File.Exists(targetPath) && File.ReadAllText(targetPath) == content)
        {
            Console.WriteLine($"  {displayName}: already up to date ({targetPath})");
            return;
        }

        Directory.CreateDirectory(Path.GetDirectoryName(targetPath)!);
        File.WriteAllText(targetPath, content);
        Console.WriteLine($"  {displayName}: installed ({targetPath})");
    }

    private static string Home => Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

    private static string? LoadEmbeddedResource(string resourceName)
    {
        var assembly = Assembly.GetExecutingAssembly();
        using var stream = assembly.GetManifestResourceStream(resourceName);
        if (stream == null) return null;
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }
}
