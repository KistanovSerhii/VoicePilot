using VoicePilot.App.Services;

namespace VoicePilot.App.Messaging;

public record SystemCommandTriggeredMessage(SystemCommandType CommandType);
