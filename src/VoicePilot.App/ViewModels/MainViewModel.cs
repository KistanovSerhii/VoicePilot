using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using CommunityToolkit.Mvvm.Messaging;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Win32;
using VoicePilot.App.Configuration;
using VoicePilot.App.Messaging;
using VoicePilot.App.Modules;
using VoicePilot.App.Resources;
using VoicePilot.App.Services;
using VoicePilot.App.Windows;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using SaveFileDialog = Microsoft.Win32.SaveFileDialog;
using Application = System.Windows.Application;

namespace VoicePilot.App.ViewModels;

public partial class MainViewModel : ObservableRecipient,
    IRecipient<RecognitionLogMessage>,
    IRecipient<RecognitionStateChangedMessage>,
    IRecipient<ModuleCatalogChangedMessage>,
    IRecipient<CommandExecutionStartedMessage>,
    IRecipient<CommandExecutedMessage>,
    IRecipient<CommandRejectedMessage>,
    IRecipient<SystemCommandTriggeredMessage>
{
    private readonly ModuleManager _moduleManager;
    private readonly SystemCommandService _systemCommandService;
    private readonly ResourceRegistry _resourceRegistry;
    private readonly IOptions<PathOptions> _pathOptions;
    private readonly ILogger<MainViewModel> _logger;
    private readonly ILoggerFactory _loggerFactory;
    private bool _modulesLoaded;
    private bool _commandExecuting;

    [ObservableProperty]
    private AssistantState _assistantState = AssistantState.Idle;

    [ObservableProperty]
    private string _assistantStateLabel = "Тишина";

    [ObservableProperty]
    private string _statusMessage = "Ожидание слова активации";

    [ObservableProperty]
    private string? _mutedUntilDescription;

    [ObservableProperty]
    private CommandListItemViewModel? _selectedCommandItem;

    [ObservableProperty]
    private string? _selectedCommandPhrases;

    [ObservableProperty]
    private string _listeningIndicatorImage = "pack://application:,,,/microphone.png";

    public ObservableCollection<LogEntryViewModel> Logs { get; } = new();

    public ObservableCollection<SystemCommandEntryViewModel> SystemCommands { get; } = new();

    public ObservableCollection<CommandListItemViewModel> CommandItems { get; } = new();

    public IAsyncRelayCommand LoadModuleCommand { get; }

    public IAsyncRelayCommand ExportModuleCommand { get; }

    public IAsyncRelayCommand ReloadModulesCommand { get; }

    public IRelayCommand OpenSystemCommandSettingsCommand { get; }

    public IRelayCommand OpenResourceManagerCommand { get; }

    public IRelayCommand OpenCommandBuilderCommand { get; }

    public MainViewModel(
        ModuleManager moduleManager,
        SystemCommandService systemCommandService,
        ResourceRegistry resourceRegistry,
        IOptions<PathOptions> pathOptions,
        ILogger<MainViewModel> logger,
        ILoggerFactory loggerFactory,
        IMessenger messenger)
        : base(messenger)
    {
        _moduleManager = moduleManager;
        _systemCommandService = systemCommandService;
                _resourceRegistry = resourceRegistry;
        _pathOptions = pathOptions;
        _logger = logger;
        _loggerFactory = loggerFactory;

        LoadModuleCommand = new AsyncRelayCommand(ImportModuleAsync);
        ExportModuleCommand = new AsyncRelayCommand(ExportSelectedModuleAsync);
        ReloadModulesCommand = new AsyncRelayCommand(ReloadModulesAsync);
        OpenSystemCommandSettingsCommand = new RelayCommand(OpenSystemCommandSettings);
        OpenResourceManagerCommand = new RelayCommand(OpenResourceManager);
        OpenCommandBuilderCommand = new RelayCommand(OpenNewCommandBuilder);

        BuildSystemCommandList();

        IsActive = true;
        UpdateIndicatorForState();
    }

    partial void OnSelectedCommandItemChanged(CommandListItemViewModel? value)
    {
        SelectedCommandPhrases = value?.Phrases;
    }

    public void Receive(RecognitionLogMessage message)
    {
        App.Current.Dispatcher.Invoke(() =>
        {
            Logs.Insert(0, new LogEntryViewModel(message.Level, message.Message));
            const int maxEntries = 200;
            while (Logs.Count > maxEntries)
            {
                Logs.RemoveAt(Logs.Count - 1);
            }
        });
    }

    public void Receive(RecognitionStateChangedMessage message)
    {
        AssistantState = message.State;

        switch (message.State)
        {
            case AssistantState.Idle:
                AssistantStateLabel = "Тишина";
                StatusMessage = "Ожидание слова активации";
                MutedUntilDescription = null;
                break;
            case AssistantState.AwaitingCommand:
                AssistantStateLabel = "Жду команду";
                StatusMessage = "Произнесите голосовую команду";
                break;
            case AssistantState.Dictation:
                AssistantStateLabel = "Голосовой ввод";
                StatusMessage = "Диктовка активна";
                break;
            case AssistantState.Muted:
                AssistantStateLabel = "Режим тишины";
                StatusMessage = "Ассистент временно не реагирует";
                MutedUntilDescription = message.Until?.ToLocalTime().ToString("t");
                break;
            case AssistantState.Listening:
                AssistantStateLabel = "Слушаю";
                StatusMessage = "Запоминаю следующую команду";
                break;
            case AssistantState.Error:
                AssistantStateLabel = "Ошибка";
                StatusMessage = "Проверьте настройки микрофона";
                break;
        }        UpdateIndicatorForState();
    }

    public void Receive(ModuleCatalogChangedMessage message)
    {
        App.Current.Dispatcher.Invoke(() =>
        {
            CommandItems.Clear();
            foreach (var module in message.Modules)
            {
                foreach (var command in module.Commands)
                {
                    CommandItems.Add(new CommandListItemViewModel(module, command));
                }
            }
        });
    }

    public void Receive(CommandExecutedMessage message)
    {
        _commandExecuting = false;
        UpdateIndicatorForState();
        Receive(new RecognitionLogMessage(LogLevel.Information,
            $"Выполнена команда \"{message.Command.Name}\" ({message.Module.Name})"));
    }

    public void Receive(CommandRejectedMessage message)
    {
        Receive(new RecognitionLogMessage(LogLevel.Warning,
            $"Команда не найдена: {message.RecognisedPhrase}"));
        _commandExecuting = false;
        UpdateIndicatorForState();
    }

    public void Receive(SystemCommandTriggeredMessage message)
    {
        App.Current.Dispatcher.Invoke(() => HighlightSystemCommands(message.CommandType));
    }

    public void Receive(CommandExecutionStartedMessage message)
    {
        _commandExecuting = true;
        ListeningIndicatorImage = "pack://application:,,,/flow.png";
    }

    private void BuildSystemCommandList()
    {
        SystemCommands.Clear();
        SystemCommands.Add(new SystemCommandEntryViewModel(
            "Активация",
            string.Join(", ", _systemCommandService.ActivationKeywords),
            SystemCommandType.Activation));

        SystemCommands.Add(new SystemCommandEntryViewModel(
            "Режим тишины",
            string.Join(", ", _systemCommandService.SilenceKeywords),
            SystemCommandType.Silence));


        SystemCommands.Add(new SystemCommandEntryViewModel(
            "Завершение",
            string.Join(", ", _systemCommandService.ExitKeywords),
            SystemCommandType.Exit));

        HighlightSystemCommands(SystemCommandType.Activation);
    }

    private void HighlightSystemCommands(SystemCommandType? activeType)
    {
        foreach (var entry in SystemCommands)
        {
            entry.IsCurrent = activeType.HasValue && entry.Type == activeType.Value;
        }
    }

    private void RefreshCommandItems(IEnumerable<CommandModule> modules)
    {
        App.Current.Dispatcher.Invoke(() =>
        {
            CommandItems.Clear();
            foreach (var module in modules)
            {
                foreach (var command in module.Commands)
                {
                    CommandItems.Add(new CommandListItemViewModel(module, command));
                }
            }
        });
    }

    public async Task EnsureModulesLoadedAsync()
    {
        if (_modulesLoaded)
        {
            RefreshCommandItems(_moduleManager.Modules);
            return;
        }

        await _moduleManager.InitialiseAsync(CancellationToken.None).ConfigureAwait(false);
        RefreshCommandItems(_moduleManager.Modules);
        _modulesLoaded = true;
    }


    private void OpenNewCommandBuilder()
    {
        ShowCommandBuilder(CommandBuilderContext.CreateNew());
    }

    public void OpenCommandEditor(CommandListItemViewModel item)
    {
        if (string.IsNullOrWhiteSpace(item.ManifestPath) || !File.Exists(item.ManifestPath))
        {
            Receive(new RecognitionLogMessage(LogLevel.Warning,
                $"Манифест команды {item.ModuleName} не найден."));
            return;
        }

        ShowCommandBuilder(CommandBuilderContext.ForExisting(item));
    }

    private void ShowCommandBuilder(CommandBuilderContext context)
    {
        var window = new CommandBuilderWindow(this, context)
        {
            Owner = App.Current?.MainWindow
        };

        window.ShowDialog();
    }

    private async Task ImportModuleAsync()
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Пакеты команд (*.zip)|*.zip|Все файлы (*.*)|*.*",
            Title = "Выберите пакет команды"
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        try
        {
            var module = await _moduleManager.ImportModuleAsync(dialog.FileName, CancellationToken.None);
            if (module is not null)
            {
                Receive(new RecognitionLogMessage(LogLevel.Information,
                    $"Команда {module.Name} ({module.Version}) установлена."));
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Не удалось загрузить модуль команд.");
            Receive(new RecognitionLogMessage(LogLevel.Error, $"Ошибка загрузки: {ex.Message}"));
        }
    }

    private Task ExportSelectedModuleAsync()
    {
        if (SelectedCommandItem is null)
        {
            Receive(new RecognitionLogMessage(LogLevel.Warning, "Выберите команду для выгрузки."));
            return Task.CompletedTask;
        }

        var moduleDirectory = SelectedCommandItem.Module.BaseDirectory;
        if (string.IsNullOrWhiteSpace(moduleDirectory) || !Directory.Exists(moduleDirectory))
        {
            Receive(new RecognitionLogMessage(LogLevel.Warning,
                $"Каталог модуля для {SelectedCommandItem.ModuleName} не найден."));
            return Task.CompletedTask;
        }

        var dialog = new SaveFileDialog
        {
            Filter = "Архивы (*.zip)|*.zip",
            FileName = $"{SanitiseSegment(SelectedCommandItem.ModuleName)}-{SelectedCommandItem.Module.Version}.zip"
        };

        if (dialog.ShowDialog() != true)
        {
            return Task.CompletedTask;
        }

        try
        {
            if (File.Exists(dialog.FileName))
            {
                File.Delete(dialog.FileName);
            }

            ZipFile.CreateFromDirectory(moduleDirectory, dialog.FileName, CompressionLevel.Optimal, includeBaseDirectory: false);
            Receive(new RecognitionLogMessage(LogLevel.Information,
                $"Модуль выгружен в {dialog.FileName}."));
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Не удалось выгрузить модуль.");
            Receive(new RecognitionLogMessage(LogLevel.Error, $"Ошибка выгрузки: {ex.Message}"));
        }

        return Task.CompletedTask;
    }
    private async Task ReloadModulesAsync()
    {
        try
        {
            await _moduleManager.InitialiseAsync(CancellationToken.None);
            Receive(new RecognitionLogMessage(LogLevel.Information, "Список команд обновлён."));
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Не удалось обновить модули.");
            Receive(new RecognitionLogMessage(LogLevel.Error, $"Ошибка обновления модулей: {ex.Message}"));
        }
    }

    private void OpenSystemCommandSettings()
    {
        var window = new SystemCommandSettingsWindow(new SystemCommandSettingsViewModel(_systemCommandService))
        {
            Owner = Application.Current?.MainWindow
        };

        var result = window.ShowDialog();
        if (result == true)
        {
            BuildSystemCommandList();
            Receive(new RecognitionLogMessage(LogLevel.Information, "Системные команды обновлены."));
        }
    }


    private void OpenResourceManager()
    {
        var viewModel = new ResourceManagerViewModel(
            _resourceRegistry,
            _loggerFactory.CreateLogger<ResourceManagerViewModel>());

        var window = new ResourceManagerWindow(viewModel)
        {
            Owner = Application.Current?.MainWindow
        };

        if (window.ShowDialog() == true)
        {
            Receive(new RecognitionLogMessage(LogLevel.Information, "Список ресурсов обновлён."));
        }
    }

    internal async Task CreateQuickCommandAsync(string commandName, string activationPhrase, string executablePath, string? arguments, string? workingDirectory)
    {
        await CreateQuickCommandAsync(commandName, activationPhrase, "runProcess", executablePath, arguments, workingDirectory);
    }

    internal async Task CreateQuickCommandAsync(string commandName, string activationPhrase, string actionType, string executablePath, string? arguments, string? workingDirectory)
    {
        var displayName = string.IsNullOrWhiteSpace(commandName) ? "Команда" : commandName.Trim();
        var sanitizedName = SanitiseSegment(displayName);
        var phrases = SplitPhrases(activationPhrase, displayName);

        // Check for name conflicts
        var moduleId = $"user.{sanitizedName}";
        var moduleDirectory = Path.Combine(_pathOptions.Value.ModulesDirectory, moduleId);
        
        if (Directory.Exists(moduleDirectory))
        {
            throw new InvalidOperationException($"Команда с именем \"{displayName}\" уже существует. Пожалуйста, выберите другое имя.");
        }
        
        Directory.CreateDirectory(moduleDirectory);

        JsonObject parameters;
        if (string.Equals(actionType, "runProcess", StringComparison.OrdinalIgnoreCase))
        {
            parameters = new JsonObject
            {
                ["path"] = executablePath,
                ["arguments"] = string.IsNullOrWhiteSpace(arguments) ? null : arguments,
                ["workingDirectory"] = string.IsNullOrWhiteSpace(workingDirectory) ? null : workingDirectory
            };
        }
        else
        {
            // For other action types, parse arguments as key-value pairs or JSON
            parameters = new JsonObject();
            if (!string.IsNullOrWhiteSpace(arguments))
            {
                try
                {
                    parameters = BuildUpdatedParameters(null, arguments);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Не удалось разобрать параметры для действия {ActionType}: {Arguments}", actionType, arguments);
                }
            }
        }

        var manifest = new
        {
            id = moduleId,
            name = displayName,
            version = "1.0.0",
            culture = "ru-RU",
            author = "Пользователь",
            description = $"Команда создана {DateTime.Now:dd.MM.yyyy}",
            commands = new[]
            {
                new
                {
                    id = Guid.NewGuid().ToString("N"),
                    name = displayName,
                    description = $"Команда {displayName}",
                    phrases,
                    actions = new[]
                    {
                        new
                        {
                            type = actionType,
                            parameters
                        }
                    }
                }
            }
        };

        var options = new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        var manifestPath = Path.Combine(moduleDirectory, "manifest.json");
        await File.WriteAllTextAsync(manifestPath, JsonSerializer.Serialize(manifest, options), Encoding.UTF8);

        await _moduleManager.InitialiseAsync(CancellationToken.None);
        Receive(new RecognitionLogMessage(LogLevel.Information,
            $"Создана команда \"{displayName}\" (фразы: {string.Join(", ", phrases)})."));
    }

    internal async Task UpdateCommandAsync(CommandBuilderContext context, string commandName, string activationPhrase, string executablePath, string? arguments, string? workingDirectory)
    {
        if (context.SourceItem is null)
        {
            throw new InvalidOperationException("Контекст команды не определён");
        }

        var manifestPath = context.SourceItem.ManifestPath;
        if (!File.Exists(manifestPath))
        {
            throw new FileNotFoundException("Файл manifest.json не найден.", manifestPath);
        }

        var json = await File.ReadAllTextAsync(manifestPath, Encoding.UTF8);
        var root = JsonNode.Parse(json)?.AsObject() ?? throw new InvalidDataException("Неверный формат manifest.json");

        root["name"] = commandName;
        root["description"] = $"Команда обновлена {DateTime.Now:dd.MM.yyyy}";

        if (root["commands"] is not JsonArray commandsArray)
        {
            commandsArray = new JsonArray();
            root["commands"] = commandsArray;
        }

        var commandNode = commandsArray
            .OfType<JsonObject>()
            .FirstOrDefault(node => node["id"]?.GetValue<string>() == context.SourceItem.Command.Id)
            ?? new JsonObject();

        if (!commandsArray.Contains(commandNode))
        {
            commandsArray.Add(commandNode);
        }

        commandNode["id"] = context.SourceItem.Command.Id;
        commandNode["name"] = commandName;
        commandNode["description"] = $"Команда обновлена {DateTime.Now:dd.MM.yyyy}";

        var phrasesArray = new JsonArray();
        foreach (var phrase in SplitPhrases(activationPhrase, commandName))
        {
            phrasesArray.Add(phrase);
        }

        commandNode["phrases"] = phrasesArray;

        var existingActions = commandNode["actions"] as JsonArray ?? new JsonArray();
        var firstAction = existingActions
            .OfType<JsonObject>()
            .FirstOrDefault();
        var existingActionType = firstAction?["type"]?
            .GetValue<string>();

        if (string.IsNullOrWhiteSpace(existingActionType) ||
            string.Equals(existingActionType, "runProcess", StringComparison.OrdinalIgnoreCase))
        {
            if (string.IsNullOrWhiteSpace(executablePath))
            {
                throw new InvalidOperationException("Путь к исполняемому файлу не может быть пустым для запуска процесса.");
            }

            var parameters = new JsonObject
            {
                ["path"] = executablePath,
                ["arguments"] = arguments,
                ["workingDirectory"] = workingDirectory
            };

            var action = new JsonObject
            {
                ["type"] = "runProcess",
                ["parameters"] = parameters
            };

            commandNode["actions"] = new JsonArray(action);
        }
        else
        {
            if (firstAction is not null && !string.IsNullOrWhiteSpace(arguments))
            {
                var existingParameters = firstAction["parameters"] as JsonObject;
                var updatedParameters = BuildUpdatedParameters(existingParameters, arguments);
                firstAction["parameters"] = updatedParameters;
            }

            commandNode["actions"] = existingActions;
        }

        var options = new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        await File.WriteAllTextAsync(manifestPath, root.ToJsonString(options), Encoding.UTF8);

        await _moduleManager.InitialiseAsync(CancellationToken.None);
        Receive(new RecognitionLogMessage(LogLevel.Information,
            $"Команда \"{commandName}\" сохранена."));
    }

    internal async Task UpdateCommandAsync(CommandBuilderContext context, string commandName, string activationPhrase, string actionType, string executablePath, string? arguments, string? workingDirectory)
    {
        if (context.SourceItem is null)
        {
            throw new InvalidOperationException("Контекст команды не определён");
        }

        var manifestPath = context.SourceItem.ManifestPath;
        if (!File.Exists(manifestPath))
        {
            throw new FileNotFoundException("Файл manifest.json не найден.", manifestPath);
        }

        // Check for name conflicts (only if name is being changed)
        var currentModuleName = context.SourceItem.Module.Id;
        var newSanitizedName = SanitiseSegment(commandName);
        var newModuleId = $"user.{newSanitizedName}";
        
        // Only check for conflicts if the name is actually changing
        if (!string.Equals(currentModuleName, newModuleId, StringComparison.OrdinalIgnoreCase))
        {
            var newModuleDirectory = Path.Combine(_pathOptions.Value.ModulesDirectory, newModuleId);
            if (Directory.Exists(newModuleDirectory))
            {
                throw new InvalidOperationException($"Команда с именем \"{commandName}\" уже существует. Пожалуйста, выберите другое имя.");
            }
        }

        var json = await File.ReadAllTextAsync(manifestPath, Encoding.UTF8);
        var root = JsonNode.Parse(json)?.AsObject() ?? throw new InvalidDataException("Неверный формат manifest.json");

        root["name"] = commandName;
        root["description"] = $"Команда обновлена {DateTime.Now:dd.MM.yyyy}";

        if (root["commands"] is not JsonArray commandsArray)
        {
            commandsArray = new JsonArray();
            root["commands"] = commandsArray;
        }

        var commandNode = commandsArray
            .OfType<JsonObject>()
            .FirstOrDefault(node => node["id"]?.GetValue<string>() == context.SourceItem.Command.Id)
            ?? new JsonObject();

        if (!commandsArray.Contains(commandNode))
        {
            commandsArray.Add(commandNode);
        }

        commandNode["id"] = context.SourceItem.Command.Id;
        commandNode["name"] = commandName;
        commandNode["description"] = $"Команда обновлена {DateTime.Now:dd.MM.yyyy}";

        var phrasesArray = new JsonArray();
        foreach (var phrase in SplitPhrases(activationPhrase, commandName))
        {
            phrasesArray.Add(phrase);
        }

        commandNode["phrases"] = phrasesArray;

        JsonObject parameters;
        if (string.Equals(actionType, "runProcess", StringComparison.OrdinalIgnoreCase))
        {
            if (string.IsNullOrWhiteSpace(executablePath))
            {
                throw new InvalidOperationException("Путь к исполняемому файлу не может быть пустым для запуска процесса.");
            }
            
            parameters = new JsonObject
            {
                ["path"] = executablePath,
                ["arguments"] = arguments,
                ["workingDirectory"] = workingDirectory
            };
        }
        else
        {
            // For other action types, parse arguments as key-value pairs or JSON
            parameters = new JsonObject();
            if (!string.IsNullOrWhiteSpace(arguments))
            {
                try
                {
                    parameters = BuildUpdatedParameters(null, arguments);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Не удалось разобрать параметры для действия {ActionType}: {Arguments}", actionType, arguments);
                }
            }
        }

        var action = new JsonObject
        {
            ["type"] = actionType,
            ["parameters"] = parameters
        };

        commandNode["actions"] = new JsonArray(action);

        var options = new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        await File.WriteAllTextAsync(manifestPath, root.ToJsonString(options), Encoding.UTF8);

        await _moduleManager.InitialiseAsync(CancellationToken.None);
        Receive(new RecognitionLogMessage(LogLevel.Information,
            $"Команда \"{commandName}\" сохранена."));
    }

    internal async Task<bool> ImportModuleFromArchiveAsync(string archivePath)
    {
        try
        {
            var module = await _moduleManager.ImportModuleAsync(archivePath, CancellationToken.None);
            if (module is not null)
            {
                Receive(new RecognitionLogMessage(LogLevel.Information,
                    $"Команда {module.Name} ({module.Version}) установлена."));
                return true;
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Не удалось загрузить модуль команд из {Path}.", archivePath);
            Receive(new RecognitionLogMessage(LogLevel.Error, $"Ошибка загрузки: {ex.Message}"));
            throw;
        }

        return false;
    }

    internal async Task ExportCommandAsync(CommandBuilderContext context, string? destinationPath = null)
    {
        if (context.SourceItem is null)
        {
            throw new InvalidOperationException("Невозможно выгрузить команду без исходного модуля");
        }

        var moduleDirectory = context.SourceItem.Module.BaseDirectory;
        if (string.IsNullOrWhiteSpace(moduleDirectory) || !Directory.Exists(moduleDirectory))
        {
            throw new DirectoryNotFoundException("Каталог модуля не найден");
        }

        var targetPath = destinationPath;
        if (string.IsNullOrWhiteSpace(targetPath))
        {
            var dialog = new SaveFileDialog
            {
                Filter = "Архивы (*.zip)|*.zip",
                FileName = $"{SanitiseSegment(context.SourceItem.ModuleName)}-{context.SourceItem.Module.Version}.zip"
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            targetPath = dialog.FileName;
        }

        if (File.Exists(targetPath))
        {
            File.Delete(targetPath);
        }

        ZipFile.CreateFromDirectory(moduleDirectory, targetPath, CompressionLevel.Optimal, includeBaseDirectory: false);
        Receive(new RecognitionLogMessage(LogLevel.Information,
            $"Модуль сохранён в {targetPath}."));
        await Task.CompletedTask;
    }

    private void UpdateIndicatorForState()
    {
        if (_commandExecuting)
        {
            ListeningIndicatorImage = "pack://application:,,,/flow.png";
            return;
        }

        ListeningIndicatorImage = AssistantState switch
        {
            AssistantState.AwaitingCommand => "pack://application:,,,/microphonelisten.png",
            AssistantState.Listening => "pack://application:,,,/microphonelisten.png",
            AssistantState.Dictation => "pack://application:,,,/flow.png",
            AssistantState.Muted => "pack://application:,,,/microphone.png",
            AssistantState.Error => "pack://application:,,,/microphone.png",
            _ => "pack://application:,,,/microphone.png"
        };
    }


    internal async Task DeleteCommandAsync(CommandBuilderContext context)
    {
        if (context.SourceItem is null)
        {
            throw new InvalidOperationException("Невозможно удалить команду без исходного описания.");
        }

        var manifestPath = context.SourceItem.ManifestPath;
        if (!File.Exists(manifestPath))
        {
            throw new FileNotFoundException("Файл manifest.json не найден.", manifestPath);
        }

        var json = await File.ReadAllTextAsync(manifestPath, Encoding.UTF8);
        var root = JsonNode.Parse(json)?.AsObject() ?? throw new InvalidDataException("Некорректный формат manifest.json");

        if (root["commands"] is not JsonArray commandsArray)
        {
            return;
        }

        var commandNode = commandsArray
            .OfType<JsonObject>()
            .FirstOrDefault(node => node["id"]?.GetValue<string>() == context.SourceItem.Command.Id);

        if (commandNode is null)
        {
            return;
        }

        commandsArray.Remove(commandNode);

        var options = new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        await File.WriteAllTextAsync(manifestPath, root.ToJsonString(options), Encoding.UTF8);
        await _moduleManager.InitialiseAsync(CancellationToken.None);

        var message = string.Format(CultureInfo.CurrentCulture, "Команда \"{0}\" удалена.", context.SourceItem.Command.Name);
        Receive(new RecognitionLogMessage(LogLevel.Information, message));
    }

private static JsonObject BuildUpdatedParameters(JsonObject? existingParameters, string overridesText)
    {
        var trimmed = overridesText.Trim();
        var baseObject = CloneJsonObject(existingParameters);

        if (string.IsNullOrWhiteSpace(trimmed))
        {
            return baseObject;
        }

        JsonObject? parsedObject = null;
        try
        {
            parsedObject = JsonNode.Parse(trimmed)?.AsObject();
        }
        catch (JsonException)
        {
            parsedObject = null;
        }

        if (parsedObject is not null)
        {
            return parsedObject;
        }

        var segments = trimmed.Split(new[] { '\r', '\n', ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (var rawSegment in segments)
        {
            var segment = rawSegment.Trim();
            if (segment.Length == 0)
            {
                continue;
            }

            var separatorIndex = segment.IndexOf('=');
            if (separatorIndex < 0)
            {
                separatorIndex = segment.IndexOf(':');
            }

            if (separatorIndex <= 0)
            {
                throw new InvalidOperationException($"Не удалось разобрать параметр \"{segment}\". Используйте формат ключ=значение или JSON.");
            }

            var key = segment[..separatorIndex].Trim();
            var valueText = segment[(separatorIndex + 1)..].Trim();
            if (string.IsNullOrEmpty(key))
            {
                throw new InvalidOperationException("Имя параметра не может быть пустым.");
            }

            baseObject[key] = ParseParameterValue(valueText);
        }

        return baseObject;
    }

    private static JsonNode ParseParameterValue(string valueText)
    {
        if (bool.TryParse(valueText, out var boolValue))
        {
            return JsonValue.Create(boolValue)!;
        }

        if (int.TryParse(valueText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var intValue))
        {
            return JsonValue.Create(intValue)!;
        }

        if (double.TryParse(valueText, NumberStyles.Float, CultureInfo.InvariantCulture, out var doubleValue))
        {
            return JsonValue.Create(doubleValue)!;
        }

        if ((valueText.StartsWith("{") && valueText.EndsWith("}")) ||
            (valueText.StartsWith("[") && valueText.EndsWith("]")))
        {
            try
            {
                var node = JsonNode.Parse(valueText);
                if (node is null)
                {
                    throw new InvalidOperationException("Не удалось разобрать вложенный JSON параметров.");
                }

                return node;
            }
            catch (JsonException ex)
            {
                throw new InvalidOperationException($"Не удалось разобрать вложенный JSON: {ex.Message}", ex);
            }
        }

        if (valueText.StartsWith('"') && valueText.EndsWith('"') && valueText.Length >= 2)
        {
            valueText = valueText.Substring(1, valueText.Length - 2);
        }

        return JsonValue.Create(valueText)!;
    }

    private static JsonObject CloneJsonObject(JsonObject? source)
    {
        if (source is null)
        {
            return new JsonObject();
        }

        return JsonNode.Parse(source.ToJsonString(new JsonSerializerOptions
        {
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        }))?.AsObject() ?? new JsonObject();
    }

    private static string[] SplitPhrases(string? input, string fallback)
    {
        var phrases = input?
            .Split(',', StringSplitOptions.RemoveEmptyEntries)
            .Select(p => p.Trim())
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        return phrases is { Length: > 0 } ? phrases : new[] { fallback };
    }

    private static string SanitiseSegment(string value)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var cleaned = new string(value
            .Select(ch => invalid.Contains(ch) ? '_' : ch)
            .ToArray()).Trim('_');

        return string.IsNullOrWhiteSpace(cleaned) ? "module" : cleaned.ToLowerInvariant();
    }
}












