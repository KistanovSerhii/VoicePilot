using System;
using System.IO;
using System.Media;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using CommunityToolkit.Mvvm.Messaging;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using NAudio.Wave;
using Vosk;
using System.Windows;
using VoicePilot.App.Configuration;
using VoicePilot.App.Input;
using VoicePilot.App.Messaging;
using VoicePilot.App.Modules;
using VoicePilot.App.Services;
using Application = System.Windows.Application;
using System.Text.Json.Nodes;

namespace VoicePilot.App.Recognition;

public class SpeechRecognitionHostedService : BackgroundService
{
    private readonly ModuleManager _moduleManager;
    private readonly CommandRouter _commandRouter;
    private readonly ActionExecutor _actionExecutor;
    private readonly SystemCommandService _systemCommandService;
    private readonly SpeechSynthesisService _speechSynthesisService;
    private readonly IOptions<SpeechOptions> _speechOptions;
    private readonly ILogger<SpeechRecognitionHostedService> _logger;
    private readonly IMessenger _messenger;
    private readonly KeyboardController _keyboardController;

    private Model? _model;
    private VoskRecognizer? _recognizer;
    private WaveInEvent? _waveIn;

    private AssistantState _state = AssistantState.Idle;
    private DateTimeOffset? _activationExpiresAt;
    private DateTimeOffset? _mutedUntil;
    private DictationSession? _dictationSession;
    private bool _activationSticky;

    private readonly SemaphoreSlim _commandSemaphore = new(1, 1);

    public SpeechRecognitionHostedService(
        ModuleManager moduleManager,
        CommandRouter commandRouter,
        ActionExecutor actionExecutor,
        SystemCommandService systemCommandService,
        SpeechSynthesisService speechSynthesisService,
        KeyboardController keyboardController,
        IOptions<SpeechOptions> speechOptions,
        ILogger<SpeechRecognitionHostedService> logger,
        IMessenger messenger)
    {
        _moduleManager = moduleManager;
        _commandRouter = commandRouter;
        _actionExecutor = actionExecutor;
        _systemCommandService = systemCommandService;
        _speechSynthesisService = speechSynthesisService;
        _keyboardController = keyboardController;
        _speechOptions = speechOptions;
        _logger = logger;
        _messenger = messenger;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        try
        {
            await _moduleManager.InitialiseAsync(stoppingToken).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to initialise module manager.");
        }

        Vosk.Vosk.SetLogLevel(-1);

        if (!TryInitialiseRecognizer())
        {
            _messenger.Send(new RecognitionLogMessage(LogLevel.Error,
                "Модель распознавания речи не найдена. Скачайте модель Vosk и разместите её в каталоге models."));
            return;
        }

        if (!TryStartAudioCapture())
        {
            _messenger.Send(new RecognitionLogMessage(LogLevel.Error,
                "Не удалось получить доступ к микрофону."));
            return;
        }

        _messenger.Send(new RecognitionLogMessage(LogLevel.Information,
            "Распознавание речи запущено. Скажите слово активации чтобы начать."));

        try
        {
            await Task.Delay(Timeout.Infinite, stoppingToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            // expected on shutdown
        }
    }

    public override Task StopAsync(CancellationToken cancellationToken)
    {
        _waveIn?.StopRecording();
        _waveIn?.Dispose();
        _waveIn = null;

        _recognizer?.Dispose();
        _recognizer = null;

        _model?.Dispose();
        _model = null;

        return base.StopAsync(cancellationToken);
    }

    private bool TryInitialiseRecognizer()
    {
        try
        {
            var modelPath = _speechOptions.Value.EnsureModelPathRooted(AppContext.BaseDirectory);
            if (!Directory.Exists(modelPath))
            {
                _logger.LogWarning("Vosk model directory {ModelPath} does not exist.", modelPath);
                return false;
            }

            _model = new Model(modelPath);
            _recognizer = new VoskRecognizer(_model, _speechOptions.Value.SampleRate);
            _recognizer.SetMaxAlternatives(0);
            _recognizer.SetWords(true);
            // Enable more frequent partial results
            _recognizer.SetPartialWords(true);

            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to initialise Vosk recognizer.");
            return false;
        }
    }

    private bool TryStartAudioCapture()
    {
        try
        {
            _waveIn = new WaveInEvent
            {
                WaveFormat = new WaveFormat(_speechOptions.Value.SampleRate, 16, 1),
                BufferMilliseconds = 60
            };

            _waveIn.DataAvailable += OnDataAvailable;
            _waveIn.RecordingStopped += (_, _) => _logger.LogInformation("Microphone recording stopped.");

            _waveIn.StartRecording();
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to start microphone capture.");
            return false;
        }
    }

    private void OnDataAvailable(object? sender, WaveInEventArgs e)
    {
        if (_recognizer is null)
        {
            return;
        }

        CheckStateTimeouts();

        try
        {
            if (_recognizer.AcceptWaveform(e.Buffer, e.BytesRecorded))
            {
                var resultJson = _recognizer.Result();
                var text = ExtractText(resultJson);
                if (!string.IsNullOrWhiteSpace(text))
                {
                    _ = HandleRecognisedTextAsync(text);
                }
            }
            else if (_speechOptions.Value.EnablePartialResults)
            {
                var partialJson = _recognizer.PartialResult();
                // Log the JSON structure for debugging
                _logger.LogDebug("Partial JSON: {PartialJson}", partialJson);
                
                var text = ExtractText(partialJson);
                if (!string.IsNullOrWhiteSpace(text))
                {
                    if (_state == AssistantState.Dictation)
                    {
                        _ = HandleDictationPartialAsync(partialJson);
                    }
                    else
                    {
                        _messenger.Send(new RecognitionLogMessage(LogLevel.Trace, $"Partial: {text}"));
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Recognizer processing failed.");
        }
    }

    private async Task HandleRecognisedTextAsync(string text)
    {
        try
        {
            // If dictation is active, treat any text as dictation input and ignore system commands.
            if (_state == AssistantState.Dictation)
            {
                await HandleDictationInputAsync(text).ConfigureAwait(false);
                return;
            }

            var commandType = _systemCommandService.Match(text, out var score);

            // Gate system commands behind activation: only Activation is allowed when not activated
            if (commandType != SystemCommandType.None)
            {
                if (commandType == SystemCommandType.Activation)
                {
                    await HandleSystemCommandAsync(commandType, text, score).ConfigureAwait(false);
                    return;
                }

                if (_state != AssistantState.AwaitingCommand)
                {
                    _logger.LogDebug("Ignoring system command '{Command}' because assistant is not activated.", commandType);
                    return;
                }

                await HandleSystemCommandAsync(commandType, text, score).ConfigureAwait(false);
                return;
            }

            if (_state == AssistantState.Muted)
            {
                if (_mutedUntil is null || _mutedUntil <= DateTimeOffset.UtcNow)
                {
                    SetState(AssistantState.Idle);
                }
                else
                {
                    _logger.LogDebug("Assistant muted, ignoring phrase '{Phrase}'.", text);
                    return;
                }
            }

            if (_state != AssistantState.AwaitingCommand)
            {
                _logger.LogDebug("Ignoring phrase '{Phrase}' because assistant is not activated.", text);
                return;
            }

            await ExecuteUserCommandAsync(text).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to handle recognised text '{Text}'.", text);
            _messenger.Send(new RecognitionLogMessage(LogLevel.Error, $"Ошибка обработки команды: {text}", ex));
            SetState(AssistantState.Error);
        }
    }

    private async Task HandleSystemCommandAsync(SystemCommandType type, string text, double score)
    {
        switch (type)
        {
            case SystemCommandType.Activation:
                if (_dictationSession is not null)
                {
                    await StopDictationAsync("Диктовка остановлена.", false).ConfigureAwait(false);
                }

                _mutedUntil = null;
                _activationSticky = true;
                SetState(AssistantState.AwaitingCommand);
                _activationExpiresAt = null;
                _messenger.Send(new RecognitionLogMessage(LogLevel.Information,
                    $"Слово активации: \"{text}\" (надёжность {score:P0})."));
                SystemSounds.Beep.Play();
                break;
            case SystemCommandType.Silence:
                if (_state != AssistantState.AwaitingCommand)
                {
                    _logger.LogDebug("Ignoring Silence: assistant is not activated.");
                    break;
                }

                await StopDictationAsync("Диктовка прервана.", false).ConfigureAwait(false);

                _mutedUntil = DateTimeOffset.UtcNow.AddSeconds(_speechOptions.Value.SilenceDurationSeconds);
                _activationSticky = false;
                SetState(AssistantState.Muted);
                _messenger.Send(new RecognitionStateChangedMessage(AssistantState.Muted,
                    "Ассистент временно отключён", _mutedUntil));
                await _speechSynthesisService.SpeakAsync("Отключаюсь", CancellationToken.None).ConfigureAwait(false);
                break;
            case SystemCommandType.Exit:
                if (_state != AssistantState.AwaitingCommand)
                {
                    _logger.LogDebug("Ignoring Exit: assistant is not activated.");
                    break;
                }
                await StopDictationAsync("Диктовка остановлена.", false).ConfigureAwait(false);
                _activationSticky = false;
                _messenger.Send(new RecognitionLogMessage(LogLevel.Information, "Получена команда выхода."));
                await ApplicationShutdownAsync().ConfigureAwait(false);
                break;
        }

        _messenger.Send(new SystemCommandTriggeredMessage(type));
    }

    private async Task HandleDictationInputAsync(string text)
    {
        var session = _dictationSession;
        if (session is null)
        {
            return;
        }

        var trimmed = text.Trim();
        if (string.IsNullOrWhiteSpace(trimmed))
        {
            return;
        }

        session.LastInput = DateTimeOffset.UtcNow;

        var printable = trimmed + " ";
        _keyboardController.TypeText(printable);
        session.AddCharacters(printable.Length);

        _messenger.Send(new RecognitionLogMessage(LogLevel.Information, $"Диктовка: {trimmed}"));

        if (session.MaxCharacters > 0 && session.CurrentCharacters >= session.MaxCharacters)
        {
            await StopDictationAsync("Диктовка завершена: достигнут лимит символов.", true).ConfigureAwait(false);
            return;
        }
    }

    private async Task<bool> StopDictationAsync(string message, bool speak)
    {
        if (_dictationSession is null)
        {
            return false;
        }

        var wasDictation = _state == AssistantState.Dictation;
        _dictationSession = null;
        _activationExpiresAt = null;

        _messenger.Send(new RecognitionLogMessage(LogLevel.Information, message));
        if (speak)
        {
            await _speechSynthesisService.SpeakAsync(message, CancellationToken.None).ConfigureAwait(false);
        }

        if (wasDictation)
        {
            // Instead of going back to previous state, switch to Silence mode
            _mutedUntil = DateTimeOffset.UtcNow.AddSeconds(_speechOptions.Value.SilenceDurationSeconds);
            _activationSticky = false;
            SetState(AssistantState.Muted);
            _messenger.Send(new RecognitionStateChangedMessage(AssistantState.Muted,
                "Ассистент временно отключён", _mutedUntil));
            await _speechSynthesisService.SpeakAsync("Отключаюсь", CancellationToken.None).ConfigureAwait(false);
        }

        return true;
    }

    private async Task ExecuteUserCommandAsync(string recognisedText)
    {
        await _commandSemaphore.WaitAsync().ConfigureAwait(false);
        try
        {
            if (_commandRouter.TryMatch(recognisedText, out var match))
            {
                _messenger.Send(new RecognitionLogMessage(LogLevel.Information,
                    $"Команда: {match.Command.Name} (модуль {match.Module.Name})"));

                _messenger.Send(new CommandExecutionStartedMessage(match.Module, match.Command));

                var dictationAction = match.Command.Actions
                    .OfType<CommandAction>()
                    .FirstOrDefault(action => string.Equals(action.Type, "dictation", StringComparison.OrdinalIgnoreCase));

                if (dictationAction is not null)
                {
                    var options = ParseDictationParameters(dictationAction.Parameters);
                    await StartDictationAsync(match.Module, match.Command, options).ConfigureAwait(false);
                    _messenger.Send(new CommandExecutedMessage(match.Module, match.Command, recognisedText, match.Score));
                    return;
                }

                await _actionExecutor.ExecuteAsync(match.Module, match.Command, CancellationToken.None)
                    .ConfigureAwait(false);

                _messenger.Send(new CommandExecutedMessage(match.Module, match.Command, recognisedText, match.Score));

                if (_activationSticky)
                {
                    // remain in AwaitingCommand until Silence
                }
                else
                {
                    SetState(AssistantState.Idle);
                }
            }
            else
            {
                _messenger.Send(new CommandRejectedMessage(recognisedText));
                _logger.LogInformation("Command not found for phrase '{Phrase}'.", recognisedText);
                await _speechSynthesisService.SpeakAsync("Команда не найдена", CancellationToken.None)
                    .ConfigureAwait(false);
                // Stay in current state; do not switch modes on unrecognized command
            }
        }
        finally
        {
            _commandSemaphore.Release();
        }
    }

    private void SetState(AssistantState state)
    {
        if (_state == state)
        {
            return;
        }

        _state = state;
        _messenger.Send(new RecognitionStateChangedMessage(_state));
    }

    private void CheckStateTimeouts()
    {
        if (_state == AssistantState.Dictation && _dictationSession is { } dictation)
        {
            if (DateTimeOffset.UtcNow - dictation.LastInput > dictation.SilenceTimeout)
            {
                _ = StopDictationAsync("Диктовка завершена: длительное молчание.", true);
            }
        }

        // Removed activation expiry: activation persists until Silence

        if (_state == AssistantState.Muted && _mutedUntil is not null && DateTimeOffset.UtcNow > _mutedUntil.Value)
        {
            _logger.LogInformation("Assistant resumed after mute timeout.");
            _mutedUntil = null;
            SetState(AssistantState.Idle);
        }
    }

    private static string ExtractText(string json)
    {
        try
        {
            using var document = JsonDocument.Parse(json);
            if (document.RootElement.TryGetProperty("text", out var textElement))
            {
                return textElement.GetString() ?? string.Empty;
            }
        }
        catch
        {
            // ignore parse errors
        }

        return string.Empty;
    }

    private static List<string> ExtractWords(string json)
    {
        var words = new List<string>();
        try
        {
            using var document = JsonDocument.Parse(json);
            
            // Try to get words from result with word-level information
            if (document.RootElement.TryGetProperty("result", out var wordsElement) &&
                wordsElement.ValueKind == JsonValueKind.Array)
            {
                foreach (var wordElement in wordsElement.EnumerateArray())
                {
                    if (wordElement.TryGetProperty("word", out var wordProperty))
                    {
                        var word = wordProperty.GetString();
                        if (!string.IsNullOrWhiteSpace(word))
                        {
                            words.Add(word);
                        }
                    }
                }
                return words;
            }
            
            // Fallback: try to get words from partial text by splitting
            string partialText = string.Empty;
            if (document.RootElement.TryGetProperty("partial", out var textElement))
            {
                partialText = textElement.GetString() ?? string.Empty;
            }
            else if (document.RootElement.TryGetProperty("text", out textElement))
            {
                partialText = textElement.GetString() ?? string.Empty;
            }
            
            if (!string.IsNullOrWhiteSpace(partialText))
            {
                // Simple split by spaces - in a real implementation, you might want to use a more sophisticated tokenizer
                words.AddRange(partialText.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));
            }
        }
        catch
        {
            // ignore parse errors
        }

        return words;
    }

    private static Task ApplicationShutdownAsync()
    {
        return Task.Run(() =>
        {
            Application.Current?.Dispatcher.Invoke(Application.Current.Shutdown);
        });
    }
    private Task HandleDictationPartialAsync(string json)
    {
        var session = _dictationSession;
        if (session is null)
        {
            return Task.CompletedTask;
        }

        var words = ExtractWords(json);
        if (words.Count == 0)
        {
            return Task.CompletedTask;
        }

        session.LastInput = DateTimeOffset.UtcNow;

        // Type only the new words to avoid duplicates from evolving partial results
        var previousWords = session.PartialPrintedWords ?? new List<string>();
        var newWords = words.Skip(previousWords.Count).ToList();

        if (newWords.Count > 0)
        {
            var textToType = string.Join(" ", newWords) + " ";
            _keyboardController.TypeText(textToType);
            session.AddCharacters(textToType.Length);
            session.PartialPrintedWords = words;
            _messenger.Send(new RecognitionLogMessage(LogLevel.Trace, $"Dictation partial typed: {string.Join(" ", newWords)}"));

            if (session.MaxCharacters > 0 && session.CurrentCharacters >= session.MaxCharacters)
            {
                _ = StopDictationAsync("Диктовка завершена: достигнут лимит символов.", true);
            }
        }

        return Task.CompletedTask;
    }

    private DictationOptions ParseDictationParameters(JsonObject? parameters)
    {
        var options = new DictationOptions();

        if (parameters is not null)
        {
            if (parameters.TryGetPropertyValue("maxCharacters", out var maxNode))
            {
                if (int.TryParse(maxNode?.ToString(), out var max))
                {
                    options.MaxCharacters = max;
                }
            }

            if (parameters.TryGetPropertyValue("silenceTimeoutSeconds", out var silenceNode))
            {
                if (int.TryParse(silenceNode?.ToString(), out var seconds))
                {
                    options.SilenceTimeoutSeconds = seconds;
                }
            }
        }

        return options;
    }

    // start dictation as a modal user command
    private async Task StartDictationAsync(CommandModule module, VoiceCommand command, DictationOptions options)
    {
        var silenceTimeout = Math.Max(1, options.SilenceTimeoutSeconds);
        _dictationSession = new DictationSession(options.MaxCharacters, TimeSpan.FromSeconds(silenceTimeout));
        SetState(AssistantState.Dictation);
        _activationExpiresAt = null;

        var limitDescription = options.MaxCharacters > 0
            ? $"лимит {options.MaxCharacters} символов"
            : "без ограничения по символам";

        _messenger.Send(new RecognitionLogMessage(LogLevel.Information,
            $"Режим диктовки активирован ({limitDescription}, пауза {silenceTimeout} с)."));
        SystemSounds.Beep.Play();
    }

    private sealed class DictationSession
    {
        public DictationSession(int maxCharacters, TimeSpan silenceTimeout)
        {
            MaxCharacters = maxCharacters;
            SilenceTimeout = silenceTimeout;
            LastInput = DateTimeOffset.UtcNow;
        }

        public int MaxCharacters { get; }

        public int CurrentCharacters { get; private set; }

        public TimeSpan SilenceTimeout { get; }

        public DateTimeOffset LastInput { get; set; }

        public List<string> PartialPrintedWords { get; set; } = new List<string>();

        public void AddCharacters(int count) => CurrentCharacters += count;
    }
}
