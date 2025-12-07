using System;
using System.Globalization;
using System.Speech.Synthesis;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;

namespace VoicePilot.App.Services;

public class SpeechSynthesisService : IDisposable
{
    private readonly SpeechSynthesizer _synthesizer = new();
    private readonly object _speechLock = new();
    private readonly ILogger<SpeechSynthesisService> _logger;

    public SpeechSynthesisService(ILogger<SpeechSynthesisService> logger)
    {
        _logger = logger;

        try
        {
            _synthesizer.SetOutputToDefaultAudioDevice();
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Unable to initialise speech synthesizer output.");
        }
    }

    public Task SpeakAsync(string text, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return Task.CompletedTask;
        }

        return Task.Run(() =>
        {
            try
            {
                lock (_speechLock)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    _synthesizer.Speak(text);
                }
            }
            catch (OperationCanceledException)
            {
                // ignored
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Speech synthesis failed.");
            }
        }, cancellationToken);
    }

    public void Dispose()
    {
        _synthesizer.Dispose();
    }
}
