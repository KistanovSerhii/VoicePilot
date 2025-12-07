namespace VoicePilot.App.Configuration;

public class DictationOptions
{
    public int MaxCharacters { get; set; } = 400;

    public int SilenceTimeoutSeconds { get; set; } = 4;
}
