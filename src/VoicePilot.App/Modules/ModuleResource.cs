using System;
using System.IO;

namespace VoicePilot.App.Modules;

public class ModuleResource
{
    public string Key { get; init; } = string.Empty;

    public string Path { get; init; } = string.Empty;

    public string Type { get; init; } = "Generic";

    public string ResolvePath(string moduleBaseDirectory)
    {
        if (System.IO.Path.IsPathRooted(Path))
        {
            return Environment.ExpandEnvironmentVariables(Path);
        }

        var combined = System.IO.Path.Combine(moduleBaseDirectory, Path);
        return Environment.ExpandEnvironmentVariables(System.IO.Path.GetFullPath(combined));
    }
}
