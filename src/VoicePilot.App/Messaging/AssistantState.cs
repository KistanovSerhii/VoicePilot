namespace VoicePilot.App.Messaging;

public enum AssistantState
{
    Idle,
    AwaitingCommand,
    Dictation,
    Muted,
    Listening,
    Error
}
