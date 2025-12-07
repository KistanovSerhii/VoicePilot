using System.Collections.Generic;
using System.Text.Json.Nodes;

namespace VoicePilot.App.Modules;

public class CommandAction
{
    public CommandAction(string type, JsonObject? parameters, IReadOnlyList<CommandAction>? children = null)
    {
        Type = type;
        Parameters = parameters ?? new JsonObject();
        Children = children ?? [];
    }

    public string Type { get; }

    public JsonObject Parameters { get; }

    public IReadOnlyList<CommandAction> Children { get; }
}
