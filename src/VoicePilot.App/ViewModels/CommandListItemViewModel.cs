using System;
using System.Linq;
using VoicePilot.App.Modules;

namespace VoicePilot.App.ViewModels;

public class CommandListItemViewModel
{
    public CommandListItemViewModel(CommandModule module, VoiceCommand command)
    {
        Module = module;
        Command = command;
    }

    public CommandModule Module { get; }

    public VoiceCommand Command { get; }

    public string ModuleName => Module.Name;

    public string ModuleId => Module.Id;

    public string Author => Module.Author;

    public string Version => Module.Version;

    public string CommandName => Command.Name;

    public string Phrases => string.Join(", ", Command.Phrases ?? Array.Empty<string>());

    public string ManifestPath
    {
        get
        {
            if (string.IsNullOrWhiteSpace(Module.BaseDirectory))
            {
                return string.Empty;
            }

            var jsonPath = System.IO.Path.Combine(Module.BaseDirectory, "manifest.json");
            if (System.IO.File.Exists(jsonPath))
            {
                return jsonPath;
            }

            var xmlPath = System.IO.Path.Combine(Module.BaseDirectory, "manifest.xml");
            if (System.IO.File.Exists(xmlPath))
            {
                return xmlPath;
            }

            return jsonPath;
        }
    }
}
