using System;
using System.IO;

namespace VoicePilot.App.Configuration;

public class SpeechOptions
{
    private static readonly string DefaultModelDirectory =
        Path.Combine("models", "vosk-model-small-ru-0.22");

    public string ModelPath { get; set; } = DefaultModelDirectory;

    public int SampleRate { get; set; } = 16_000;

    public double CommandMatchThreshold { get; set; } = 0.72;

    public double CommandMatchFallbackThreshold { get; set; } = 0.65;

    public int ActivationTimeoutSeconds { get; set; } = 12;

    public int SilenceDurationSeconds { get; set; } = 30;

    public bool EnablePartialResults { get; set; } = true;

    public string Culture { get; set; } = "ru-RU";

    public string EnsureModelPathRooted(string baseDirectory)
    {
        if (Path.IsPathRooted(ModelPath))
        {
            return ModelPath;
        }

        return Path.GetFullPath(Path.Combine(baseDirectory, ModelPath));
    }
}
