using System;

namespace VoicePilot.App.Messaging;

public record RecognitionStateChangedMessage(AssistantState State, string? Detail = null, DateTimeOffset? Until = null);
