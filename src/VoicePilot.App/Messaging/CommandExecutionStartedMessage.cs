using VoicePilot.App.Modules;

namespace VoicePilot.App.Messaging;

public record CommandExecutionStartedMessage(CommandModule Module, VoiceCommand Command);
