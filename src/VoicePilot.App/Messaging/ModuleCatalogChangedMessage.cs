using System.Collections.Generic;
using VoicePilot.App.Modules;

namespace VoicePilot.App.Messaging;

public record ModuleCatalogChangedMessage(IReadOnlyList<CommandModule> Modules);
