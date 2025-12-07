using System;
using System.Windows;
using VoicePilot.App.ViewModels;

namespace VoicePilot.App.Windows;

public partial class SystemCommandSettingsWindow : Window
{
    private readonly SystemCommandSettingsViewModel _viewModel;

    public SystemCommandSettingsWindow(SystemCommandSettingsViewModel viewModel)
    {
        InitializeComponent();
        _viewModel = viewModel;
        DataContext = viewModel;
        _viewModel.CloseRequested += OnCloseRequested;
        Closed += (_, _) => _viewModel.CloseRequested -= OnCloseRequested;
    }

    private void OnCloseRequested(object? sender, bool e)
    {
        DialogResult = e;
        Close();
    }
}
