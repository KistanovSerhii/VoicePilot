using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using VoicePilot.App.Configuration;
using VoicePilot.App.Services;
using MessageBox = System.Windows.MessageBox;

namespace VoicePilot.App.ViewModels;

public partial class SystemCommandSettingsViewModel : ObservableObject
{
    private readonly SystemCommandService _systemCommandService;
    private readonly ILogger<SystemCommandSettingsViewModel> _logger;

    public event EventHandler<bool>? CloseRequested;

    [ObservableProperty]
    private string _activationText = string.Empty;

    [ObservableProperty]
    private string _silenceText = string.Empty;

    [ObservableProperty]
    private string _exitText = string.Empty;

    [ObservableProperty]
    private bool _isBusy;

    [ObservableProperty]
    private string? _errorMessage;

    public SystemCommandSettingsViewModel(SystemCommandService systemCommandService, ILogger<SystemCommandSettingsViewModel>? logger = null)
    {
        _systemCommandService = systemCommandService;
        _logger = logger ?? NullLogger<SystemCommandSettingsViewModel>.Instance;

        LoadFromService();
    }

    private void LoadFromService()
    {
        var manifest = _systemCommandService.GetManifestSnapshot();
        ActivationText = string.Join(Environment.NewLine, manifest.Activation);
        SilenceText = string.Join(Environment.NewLine, manifest.Silence);
        ExitText = string.Join(Environment.NewLine, manifest.Exit);
        ErrorMessage = null;
    }

    [RelayCommand]
    private void Reset() => LoadFromService();

    [RelayCommand]
    private void Cancel() => CloseRequested?.Invoke(this, false);

    [RelayCommand]
    private void Save()
    {
        if (IsBusy)
        {
            return;
        }

        try
        {
            IsBusy = true;
            ErrorMessage = null;

            var manifest = new SystemCommandManifest
            {
                Activation = ParseInput(ActivationText),
                Silence = ParseInput(SilenceText),
                Exit = ParseInput(ExitText)
            };

            if (!Validate(manifest))
            {
                return;
            }

            _systemCommandService.SaveManifest(manifest);
            CloseRequested?.Invoke(this, true);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to save system commands.");
            ErrorMessage = $"Не удалось сохранить команды: {ex.Message}";
            MessageBox.Show(ErrorMessage, "VoicePilot", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            IsBusy = false;
        }
    }

    private static List<string> ParseInput(string? input) =>
        string.IsNullOrWhiteSpace(input)
            ? new List<string>()
            : input.Split(new[] { '\r', '\n', ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(p => p.Trim())
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

    private bool Validate(SystemCommandManifest manifest)
    {
        if (manifest.Activation.Count == 0)
        {
            ErrorMessage = "Добавьте хотя бы одно слово активации.";
            MessageBox.Show(ErrorMessage, "VoicePilot", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        if (manifest.Silence.Count == 0)
        {
            ErrorMessage = "Добавьте хотя бы одну команду тишины.";
            MessageBox.Show(ErrorMessage, "VoicePilot", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        return true;
    }
}
