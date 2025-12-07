using System;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using VoicePilot.App.ViewModels;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

namespace VoicePilot.App.Windows;

public partial class CommandBuilderWindow : Window
{
    private readonly MainViewModel _mainViewModel;
    private readonly CommandBuilderContext _context;
    private bool _phraseEdited;
    private bool _syncing;
    private string _loadedActionType = "runprocess";
    private bool _definitionLoadedFromFile;
    private string? _loadedArchivePath;
    private JsonObject? _loadedManifestRoot;
    private JsonObject? _loadedCommandNode;
    private JsonObject? _currentActionParameters;
    private bool _isUpdatingCheckbox;

    public CommandBuilderWindow(MainViewModel mainViewModel, CommandBuilderContext context)
    {
        InitializeComponent();
        _mainViewModel = mainViewModel;
        _context = context;

        ConfigureExecutionFields();

        if (_context.IsEditMode && _context.SourceItem is not null)
        {
            Title = "Редактирование команды";
            LoadButton.IsEnabled = false;
            ExportButton.IsEnabled = true;
            DeleteButton.Visibility = Visibility.Visible;
            PopulateFromCommand(_context.SourceItem);
        }
        else
        {
            Title = "Новая команда";
            ExportButton.IsEnabled = false;
            DeleteButton.Visibility = Visibility.Collapsed;
        }
    }

    private void PopulateFromCommand(CommandListItemViewModel item)
    {
        CommandNameTextBox.Text = item.Command.Name;
        ActivationPhraseTextBox.Text = string.Join(", ", item.Command.Phrases ?? Array.Empty<string>());

        var action = item.Command.Actions.FirstOrDefault();
        if (action is not null)
        {
            _loadedActionType = action.Type?.ToLowerInvariant() ?? "runprocess";
            _currentActionParameters = null;

            // Update checkbox state
            _isUpdatingCheckbox = true;
            UseResourceCheckBox.IsChecked = string.Equals(_loadedActionType, "openResource", StringComparison.OrdinalIgnoreCase);
            _isUpdatingCheckbox = false;

            if (string.Equals(_loadedActionType, "runprocess", StringComparison.OrdinalIgnoreCase))
            {
                if (action.Parameters.TryGetPropertyValue("path", out var pathNode) && pathNode is JsonValue pathValue)
                {
                    pathValue.TryGetValue(out string? path);
                    ExecutablePathTextBox.Text = path ?? string.Empty;
                }

                if (action.Parameters.TryGetPropertyValue("arguments", out var argNode) && argNode is JsonValue argValue)
                {
                    argValue.TryGetValue(out string? args);
                    ArgumentsTextBox.Text = args ?? string.Empty;
                }

                if (action.Parameters.TryGetPropertyValue("workingDirectory", out var dirNode) && dirNode is JsonValue dirValue)
                {
                    dirValue.TryGetValue(out string? dir);
                    WorkingDirectoryTextBox.Text = dir ?? string.Empty;
                }
            }
            else
            {
                ExecutablePathTextBox.Text = ResolveActionDescription(_loadedActionType);
                WorkingDirectoryTextBox.Text = string.Empty;
                _currentActionParameters = CloneParameters(action.Parameters);
                ArgumentsTextBox.Text = FormatParameters(_currentActionParameters);
            }
        }

        ConfigureExecutionFields();
        _phraseEdited = true;
    }

    private static string ResolveActionDescription(string actionType) =>
        actionType switch
        {
            "sendkeys" => "Использует действие sendKeys",
            "mousemove" => "Использует действие mouseMove",
            "mouseclick" => "Использует действие mouseClick",
            "openurl" => "Открывает ссылку",
            "openresource" => "Использует встроенное действие openResource",
            _ => "Использует встроенное действие"
        };

    private static string FormatParameters(JsonObject? parameters)
    {
        if (parameters is null || parameters.Count == 0)
        {
            return string.Empty;
        }

        return parameters.ToJsonString(new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        });
    }

    private static JsonObject CloneParameters(JsonObject parameters)
    {
        return JsonNode.Parse(parameters.ToJsonString(new JsonSerializerOptions
        {
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        }))?.AsObject() ?? new JsonObject();
    }


    private void ConfigureExecutionFields()
    {
        var isRunProcess = string.Equals(_loadedActionType, "runprocess", StringComparison.OrdinalIgnoreCase);
        var isOpenResource = string.Equals(_loadedActionType, "openresource", StringComparison.OrdinalIgnoreCase);
        var requiresExecutable = isRunProcess;

        ExecutablePathTextBox.IsEnabled = requiresExecutable;
        ExecutablePathTextBox.IsReadOnly = !requiresExecutable;
        WorkingDirectoryTextBox.IsEnabled = requiresExecutable;
        WorkingDirectoryTextBox.IsReadOnly = !requiresExecutable;
        BrowseExecutableButton.IsEnabled = requiresExecutable;
        ArgumentsTextBox.IsEnabled = true;

        if (isRunProcess)
        {
            ArgumentsLabel.Text = "Аргументы:";
            if (string.IsNullOrWhiteSpace(CommandHintTextBlock.Text))
            {
                CommandHintTextBlock.Text =
                    "Укажите исполняемый файл, аргументы командной строки и при необходимости рабочую папку.";
            }
        }
        else if (isOpenResource)
        {
            ArgumentsLabel.Text = "Параметры действия:";
            CommandHintTextBlock.Text =
                "Для открытия ресурса укажите параметр key=название_ресурса. Например: key=фильмы. Ресурсы настраиваются в меню Ресурсы.";

            if (string.IsNullOrWhiteSpace(ArgumentsTextBox.Text) && _currentActionParameters is not null)
            {
                ArgumentsTextBox.Text = FormatParameters(_currentActionParameters);
            }
        }
        else
        {
            ArgumentsLabel.Text = "Параметры действия:";
            CommandHintTextBlock.Text =
                "Измените параметры встроенного действия. Допустим формат JSON или строки вида «ключ=значение».";

            if (string.IsNullOrWhiteSpace(ArgumentsTextBox.Text) && _currentActionParameters is not null)
            {
            ArgumentsTextBox.Text = FormatParameters(_currentActionParameters);
            }
        }
    }

    private void BrowseExecutable_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Исполняемые файлы (*.exe)|*.exe|Все файлы (*.*)|*.*"
        };

        if (dialog.ShowDialog() == true)
        {
            ExecutablePathTextBox.Text = dialog.FileName;

            if (string.IsNullOrWhiteSpace(WorkingDirectoryTextBox.Text))
            {
                WorkingDirectoryTextBox.Text = Path.GetDirectoryName(dialog.FileName) ?? string.Empty;
            }
        }
    }

    
    private async void DeleteButton_Click(object sender, RoutedEventArgs e)
    {
        if (!_context.IsEditMode || _context.SourceItem is null)
        {
            return;
        }

        var commandName = _context.SourceItem.Command.Name;
        var prompt = string.Format(CultureInfo.CurrentCulture, "Удалить команду \"{0}\"?", commandName);
        var result = MessageBox.Show(this, prompt, "Подтверждение удаления", MessageBoxButton.YesNo, MessageBoxImage.Question);

        if (result != MessageBoxResult.Yes)
        {
            return;
        }

        try
        {
            await _mainViewModel.DeleteCommandAsync(_context);
            MessageBox.Show(this, "Команда удалена.", "VoicePilot", MessageBoxButton.OK, MessageBoxImage.Information);
            DialogResult = true;
        }
        catch (Exception ex)
        {
            var error = string.Format(CultureInfo.CurrentCulture, "Не удалось удалить команду: {0}", ex.Message);
            MessageBox.Show(this, error, "VoicePilot", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

private async void SaveButton_Click(object sender, RoutedEventArgs e)
    {
        var name = CommandNameTextBox.Text.Trim();
        var phrase = ActivationPhraseTextBox.Text.Trim();
        var path = ExecutablePathTextBox.Text.Trim();
        var arguments = string.IsNullOrWhiteSpace(ArgumentsTextBox.Text) ? null : ArgumentsTextBox.Text.Trim();
        var workingDirectory = string.IsNullOrWhiteSpace(WorkingDirectoryTextBox.Text) ? null : WorkingDirectoryTextBox.Text.Trim();

        if (string.IsNullOrWhiteSpace(name))
        {
            MessageBox.Show(this, "Введите название команды.", "Конструктор команд", MessageBoxButton.OK, MessageBoxImage.Warning);
            CommandNameTextBox.Focus();
            return;
        }

        if (!_context.IsEditMode && _loadedArchivePath is not null)
        {
            try
            {
                await ImportLoadedArchiveAsync(name, phrase);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Не удалось загрузить команду: {ex.Message}", "Конструктор команд", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            return;
        }

        var requiresExecutable =
            string.Equals(_loadedActionType, "runprocess", StringComparison.OrdinalIgnoreCase);

        if (requiresExecutable && string.IsNullOrWhiteSpace(path))
        {
            MessageBox.Show(this, "Укажите, что необходимо выполнить.", "Конструктор команд", MessageBoxButton.OK, MessageBoxImage.Warning);
            ExecutablePathTextBox.Focus();
            return;
        }

        try
        {
            if (_context.IsEditMode)
            {
                await _mainViewModel.UpdateCommandAsync(_context, name, phrase, _loadedActionType, requiresExecutable ? path : string.Empty, arguments, workingDirectory);
                MessageBox.Show(this, "Команда сохранена.", "VoicePilot", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                await _mainViewModel.CreateQuickCommandAsync(name, phrase, _loadedActionType, path, arguments, workingDirectory);
                MessageBox.Show(this, "Команда сохранена.", "VoicePilot", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            DialogResult = true;
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("уже существует"))
        {
            MessageBox.Show(this, ex.Message, "Конструктор команд", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Не удалось сохранить команду: {ex.Message}", "Конструктор команд", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void CommandNameTextBox_OnTextChanged(object sender, TextChangedEventArgs e)
    {
        if (_phraseEdited)
        {
            return;
        }

        _syncing = true;
        ActivationPhraseTextBox.Text = CommandNameTextBox.Text;
        ActivationPhraseTextBox.CaretIndex = ActivationPhraseTextBox.Text.Length;
        _syncing = false;
    }

    private void ActivationPhraseTextBox_OnTextChanged(object sender, TextChangedEventArgs e)
    {
        if (_syncing)
        {
            return;
        }

        _phraseEdited = !string.IsNullOrWhiteSpace(ActivationPhraseTextBox.Text) &&
                        !string.Equals(ActivationPhraseTextBox.Text, CommandNameTextBox.Text, StringComparison.Ordinal);
    }

    private void LoadButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Манифесты и архивы|manifest.json;*.json;*.zip|Все файлы (*.*)|*.*"
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        try
        {
            var data = LoadCommandFromFile(dialog.FileName);
            if (data is null)
            {
                MessageBox.Show(this, "Не удалось распознать команду в выбранном файле.", "Конструктор команд", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            CommandNameTextBox.Text = data.Value.Name;
            ActivationPhraseTextBox.Text = data.Value.Phrases;
            ExecutablePathTextBox.Text = data.Value.Path ?? string.Empty;
            ArgumentsTextBox.Text = data.Value.Arguments ?? string.Empty;
            WorkingDirectoryTextBox.Text = data.Value.WorkingDirectory ?? string.Empty;
            _loadedActionType = data.Value.ActionType;
            _currentActionParameters = data.Value.Parameters is not null ? CloneParameters(data.Value.Parameters) : null;
            
            // Update checkbox state
            _isUpdatingCheckbox = true;
            UseResourceCheckBox.IsChecked = string.Equals(_loadedActionType, "openresource", StringComparison.OrdinalIgnoreCase);
            _isUpdatingCheckbox = false;
            
            if (!string.Equals(_loadedActionType, "runprocess", StringComparison.OrdinalIgnoreCase))
            {
                ExecutablePathTextBox.Text = string.Empty;
                WorkingDirectoryTextBox.Text = string.Empty;
                ArgumentsTextBox.Text = FormatParameters(_currentActionParameters);
            }
            _definitionLoadedFromFile = true;
            _loadedArchivePath = string.Equals(Path.GetExtension(dialog.FileName), ".zip", StringComparison.OrdinalIgnoreCase)
                ? dialog.FileName
                : null;
            _loadedManifestRoot = data.Value.ManifestRoot;
            _loadedCommandNode = data.Value.CommandNode;
            ConfigureExecutionFields();
            _phraseEdited = true;
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Ошибка загрузки команды: {ex.Message}", "Конструктор команд", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private LoadedCommand? LoadCommandFromFile(string fileName)
    {
        string? manifestJson;
        if (string.Equals(Path.GetExtension(fileName), ".zip", StringComparison.OrdinalIgnoreCase))
        {
            using var archive = ZipFile.OpenRead(fileName);
            var entry = archive.GetEntry("manifest.json");
            if (entry is null)
            {
                return null;
            }

            using var stream = entry.Open();
            using var reader = new StreamReader(stream, Encoding.UTF8);
            manifestJson = reader.ReadToEnd();
        }
        else
        {
            manifestJson = File.ReadAllText(fileName, Encoding.UTF8);
        }

        var root = JsonNode.Parse(manifestJson)?.AsObject();
        if (root is null)
        {
            return null;
        }

        var commandsArray = root["commands"] as JsonArray;
        var firstCommand = commandsArray?.OfType<JsonObject>().FirstOrDefault();
        if (firstCommand is null)
        {
            return null;
        }

        var name = firstCommand["name"]?.GetValue<string>() ?? "Команда";
        var phrases = firstCommand["phrases"] is JsonArray phraseArray
            ? string.Join(", ", phraseArray.OfType<JsonValue>().Select(v => v.GetValue<string>()).Where(p => !string.IsNullOrWhiteSpace(p)).Select(p => p.Trim()))
            : name;

        if (string.IsNullOrWhiteSpace(phrases))
        {
            phrases = name;
        }

        string? path = null;
        string? args = null;
        string? dir = null;
        JsonObject? parameters = null;

        string actionType = "runprocess";

        if (firstCommand["actions"] is JsonArray actionsArray)
        {
            var action = actionsArray.OfType<JsonObject>().FirstOrDefault();
            if (action is not null)
            {
                actionType = action["type"]?.GetValue<string>()?.ToLowerInvariant() ?? actionType;
            }

            if (action?["parameters"] is JsonObject paramObject)
            {
                parameters = paramObject;
                if (string.Equals(actionType, "runprocess", StringComparison.OrdinalIgnoreCase))
                {
                    path = paramObject["path"]?.GetValue<string>();
                    args = paramObject["arguments"]?.GetValue<string>();
                    dir = paramObject["workingDirectory"]?.GetValue<string>();
                }
                else
                {
                    args = paramObject.ToJsonString(new JsonSerializerOptions
                    {
                        WriteIndented = true,
                        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                    });
                }
            }
        }

        return new LoadedCommand(name, phrases, path, args, dir, actionType, root, firstCommand, parameters);
    }

    private async void ExportButton_Click(object sender, RoutedEventArgs e)
    {
        if (!_context.IsEditMode)
        {
            MessageBox.Show(this, "Выгрузка доступна только для существующих команд.", "Конструктор команд", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        try
        {
            await _mainViewModel.ExportCommandAsync(_context);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Не удалось выгрузить модуль: {ex.Message}", "Конструктор команд", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private record struct LoadedCommand(
        string Name,
        string Phrases,
        string? Path,
        string? Arguments,
        string? WorkingDirectory,
        string ActionType,
        JsonObject ManifestRoot,
        JsonObject CommandNode,
        JsonObject? Parameters);

    private async Task ImportLoadedArchiveAsync(string commandName, string phrasesText)
    {
        if (_loadedArchivePath is null)
        {
            throw new InvalidOperationException("Архив команды не выбран.");
        }

        var tempExtract = Path.Combine(Path.GetTempPath(), "VoicePilot", "import", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempExtract);

        string preparedArchive = Path.Combine(Path.GetTempPath(), $"voicepilot-module-{Guid.NewGuid():N}.zip");

        try
        {
            ZipFile.ExtractToDirectory(_loadedArchivePath, tempExtract, overwriteFiles: true);

            var manifestPath = Directory.EnumerateFiles(tempExtract, "manifest.json", SearchOption.AllDirectories)
                .FirstOrDefault();

            if (manifestPath is null)
            {
                throw new InvalidOperationException("В архиве отсутствует manifest.json.");
            }

            var manifestRoot = JsonNode.Parse(await File.ReadAllTextAsync(manifestPath, Encoding.UTF8))?.AsObject()
                               ?? throw new InvalidDataException("Некорректный manifest.json.");

            var commandsArray = manifestRoot["commands"] as JsonArray ?? new JsonArray();
            manifestRoot["commands"] = commandsArray;

            var commandNode = commandsArray.OfType<JsonObject>().FirstOrDefault() ?? new JsonObject();
            if (!commandsArray.Contains(commandNode))
            {
                commandsArray.Add(commandNode);
            }

            commandNode["name"] = commandName;
            commandNode["phrases"] = BuildPhraseArray(phrasesText, commandName);

            if (manifestRoot.ContainsKey("name"))
            {
                manifestRoot["name"] = commandName;
            }

            await File.WriteAllTextAsync(manifestPath,
                manifestRoot.ToJsonString(new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                }),
                Encoding.UTF8);

            ZipFile.CreateFromDirectory(tempExtract, preparedArchive, CompressionLevel.Optimal, includeBaseDirectory: false);

            await _mainViewModel.ImportModuleFromArchiveAsync(preparedArchive);
            MessageBox.Show(this, "Команда установлена.", "VoicePilot", MessageBoxButton.OK, MessageBoxImage.Information);
            DialogResult = true;
        }
        finally
        {
            try
            {
                if (File.Exists(preparedArchive))
                {
                    File.Delete(preparedArchive);
                }

                if (Directory.Exists(tempExtract))
                {
                    Directory.Delete(tempExtract, recursive: true);
                }
            }
            catch
            {
                // ignore cleanup errors
            }
        }
    }

    private static JsonArray BuildPhraseArray(string phrasesText, string fallback)
    {
        var array = new JsonArray();
        var phrases = phrasesText.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        if (phrases.Length == 0)
        {
            phrases = new[] { fallback };
        }

        foreach (var phrase in phrases)
        {
            array.Add(phrase);
        }

        return array;
    }

    private void UseResourceCheckBox_OnChecked(object sender, RoutedEventArgs e)
    {
        if (_isUpdatingCheckbox)
        {
            return;
        }

        _loadedActionType = "openresource";
        ExecutablePathTextBox.Text = string.Empty;
        WorkingDirectoryTextBox.Text = string.Empty;
        
        if (string.IsNullOrWhiteSpace(ArgumentsTextBox.Text))
        {
            ArgumentsTextBox.Text = "";
        }
        
        ConfigureExecutionFields();
    }

    private void UseResourceCheckBox_OnUnchecked(object sender, RoutedEventArgs e)
    {
        if (_isUpdatingCheckbox)
        {
            return;
        }

        _loadedActionType = "runprocess";
        ExecutablePathTextBox.Text = string.Empty;
        ArgumentsTextBox.Text = string.Empty;
        ConfigureExecutionFields();
    }
}
