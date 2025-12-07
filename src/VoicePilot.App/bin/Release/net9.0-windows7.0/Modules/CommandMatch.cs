namespace VoicePilot.App.Modules;

public record CommandMatch(
    CommandModule Module,
    VoiceCommand Command,
    string MatchedPhrase,
    double Score);
