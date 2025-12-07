using System;
using System.IO;
using System.Linq;

namespace VoicePilot.App.Configuration;

public class PathOptions
{
    public string ModulesDirectory { get; set; } = "Modules";

    public string ModelsDirectory { get; set; } = "models";

    public string SystemCommandsFile { get; set; } = Path.Combine("config", "system-commands.json");

    public string ResourcesFile { get; set; } = Path.Combine("config", "resources.json");

    public string DictationSettingsFile { get; set; } = Path.Combine("config", "dictation-settings.json");

    public string ModuleStagingDirectory { get; set; } = Path.Combine("cache", "module-staging");

    public void Resolve(string baseDirectory)
    {
        var solutionRoot = FindSolutionRoot(baseDirectory);

        ModulesDirectory = ResolveWithRoot(baseDirectory, solutionRoot, ModulesDirectory, ensureDirectory: true);
        ModelsDirectory = ResolveWithRoot(baseDirectory, solutionRoot, ModelsDirectory, ensureDirectory: true);
        SystemCommandsFile = ResolveWithRoot(baseDirectory, solutionRoot, SystemCommandsFile, ensureDirectory: false);
        ResourcesFile = ResolveWithRoot(baseDirectory, solutionRoot, ResourcesFile, ensureDirectory: false);
        DictationSettingsFile = ResolveWithRoot(baseDirectory, solutionRoot, DictationSettingsFile, ensureDirectory: false);
        ModuleStagingDirectory = ResolvePath(baseDirectory, ModuleStagingDirectory, ensureDirectory: true);
    }

    private static string ResolveWithRoot(string baseDirectory, string rootDirectory, string path, bool ensureDirectory)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("Path must not be empty.", nameof(path));
        }

        if (Path.IsPathRooted(path))
        {
            return Ensure(Path.GetFullPath(path), ensureDirectory);
        }

        var candidate = Path.Combine(rootDirectory, path);
        return Ensure(Path.GetFullPath(candidate), ensureDirectory);
    }

    private static string ResolvePath(string baseDirectory, string path, bool ensureDirectory)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("Path must not be empty.", nameof(path));
        }

        var target = Path.IsPathRooted(path)
            ? Path.GetFullPath(path)
            : Path.GetFullPath(Path.Combine(baseDirectory, path));

        return Ensure(target, ensureDirectory);
    }

    private static string Ensure(string path, bool ensureDirectory)
    {
        if (ensureDirectory)
        {
            Directory.CreateDirectory(path);
        }
        else
        {
            var directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }
        }

        return path;
    }

    private static string FindSolutionRoot(string baseDirectory)
    {
        var current = new DirectoryInfo(baseDirectory);
        DirectoryInfo? modulesCandidate = null;

        while (current is not null)
        {
            var solutionFile = Path.Combine(current.FullName, "VoicePilot.sln");
            if (File.Exists(solutionFile))
            {
                return current.FullName;
            }

            var modulesFolder = Path.Combine(current.FullName, "Modules");
            if (modulesCandidate is null && Directory.Exists(modulesFolder))
            {
                modulesCandidate = current;
            }

            current = current.Parent;
        }

        return modulesCandidate?.FullName ?? baseDirectory;
    }
}

