using System;
using System.Collections.Generic;

namespace VoicePilot.App.Modules;

public class VoiceCommand
{
    public string Id { get; init; } = Guid.NewGuid().ToString("N");

    public string Name { get; init; } = string.Empty;

    public string Description { get; init; } = string.Empty;

    public IReadOnlyList<string> Phrases { get; init; } = Array.Empty<string>();

    public IReadOnlyList<CommandAction> Actions { get; init; } = Array.Empty<CommandAction>();

    public double? CustomThreshold { get; init; }

    public bool KeepAssistantActive { get; init; }
}
