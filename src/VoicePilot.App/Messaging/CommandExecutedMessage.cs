using VoicePilot.App.Modules;

namespace VoicePilot.App.Messaging;

public record CommandExecutedMessage(
    CommandModule Module,
    VoiceCommand Command,
    string RecognisedPhrase,
    double Score);
