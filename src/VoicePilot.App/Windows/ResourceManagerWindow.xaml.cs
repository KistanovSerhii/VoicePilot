using System;
using System.Windows;
using VoicePilot.App.ViewModels;

namespace VoicePilot.App.Windows;

public partial class ResourceManagerWindow : Window
{
    private readonly ResourceManagerViewModel _viewModel;

    public ResourceManagerWindow(ResourceManagerViewModel viewModel)
    {
        InitializeComponent();
        _viewModel = viewModel;
        DataContext = _viewModel;
        _viewModel.CloseRequested += OnCloseRequested;
        Closed += OnClosed;
    }

    private void OnClosed(object? sender, EventArgs e)
    {
        _viewModel.CloseRequested -= OnCloseRequested;
    }

    private void OnCloseRequested(object? sender, bool e)
    {
        DialogResult = e;
        Close();
    }
}
