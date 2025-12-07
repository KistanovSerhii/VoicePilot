using System.Collections.Generic;

namespace VoicePilot.App.Configuration;

public class SystemCommandManifest
{
    public List<string> Activation { get; init; } = new();

    public List<string> Exit { get; init; } = new();

    public List<string> Silence { get; init; } = new();
}
