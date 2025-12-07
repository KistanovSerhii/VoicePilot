using System;
using System.Collections.Generic;

namespace VoicePilot.App.Modules;

public class CommandModule
{
    public string Id { get; init; } = string.Empty;

    public string Name { get; init; } = string.Empty;

    public string Version { get; init; } = "1.0.0";

    public string Culture { get; init; } = "ru-RU";

    public string Author { get; init; } = "Unknown";

    public string Description { get; init; } = string.Empty;

    public string BaseDirectory { get; init; } = string.Empty;

    public IReadOnlyList<VoiceCommand> Commands { get; init; } = Array.Empty<VoiceCommand>();

    public IReadOnlyDictionary<string, ModuleResource> Resources { get; init; } =
        new Dictionary<string, ModuleResource>(StringComparer.OrdinalIgnoreCase);
}
