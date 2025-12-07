using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Win32;
using VoicePilot.App.Resources;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

namespace VoicePilot.App.ViewModels;

public partial class ResourceManagerViewModel : ObservableObject
{
    private readonly ResourceRegistry _resourceRegistry;
    private readonly ILogger<ResourceManagerViewModel> _logger;

    public event EventHandler<bool>? CloseRequested;

    [ObservableProperty]
    private ResourceEditorItem? _selectedResource;

    public ObservableCollection<ResourceEditorItem> Resources { get; } = new();

    public ResourceManagerViewModel(ResourceRegistry resourceRegistry, ILogger<ResourceManagerViewModel>? logger = null)
    {
        _resourceRegistry = resourceRegistry;
        _logger = logger ?? NullLogger<ResourceManagerViewModel>.Instance;

        Load();
    }

    [RelayCommand]
    private void Add()
    {
        var item = new ResourceEditorItem
        {
            Key = "новый_ресурс",
            Path = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            Type = "Folder"
        };

        Resources.Add(item);
        SelectedResource = item;
    }

    [RelayCommand]
    private void Remove()
    {
        if (SelectedResource is null)
        {
            return;
        }

        if (MessageBox.Show($"Удалить ресурс \"{SelectedResource.Key}\"?", "VoicePilot",
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
        {
            return;
        }

        Resources.Remove(SelectedResource);
        SelectedResource = null;
    }

    [RelayCommand]
    private void BrowsePath()
    {
        if (SelectedResource is null)
        {
            return;
        }

        if (string.Equals(SelectedResource.Type, "Folder", StringComparison.OrdinalIgnoreCase))
        {
            using var dialog = new System.Windows.Forms.FolderBrowserDialog
            {
                SelectedPath = Environment.ExpandEnvironmentVariables(SelectedResource.Path ?? string.Empty)
            };

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                SelectedResource.Path = dialog.SelectedPath;
            }
        }
        else
        {
            var dialog = new OpenFileDialog
            {
                FileName = Environment.ExpandEnvironmentVariables(SelectedResource.Path ?? string.Empty),
                CheckFileExists = false
            };

            if (dialog.ShowDialog() == true)
            {
                SelectedResource.Path = dialog.FileName;
            }
        }
    }

    [RelayCommand]
    private void Save()
    {
        try
        {
            var descriptors = Resources
                .Select(item => item.ToDescriptor())
                .ToList();

            if (descriptors.Count == 0)
            {
                MessageBox.Show("Добавьте хотя бы один ресурс.", "VoicePilot",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var duplicate = descriptors
                .GroupBy(r => r.Key, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault(g => g.Count() > 1);

            if (duplicate is not null)
            {
                MessageBox.Show($"Ключ \"{duplicate.Key}\" повторяется. Сделайте его уникальным.",
                    "VoicePilot", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _resourceRegistry.SaveResources(descriptors);
            CloseRequested?.Invoke(this, true);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to save resources.");
            MessageBox.Show($"Не удалось сохранить ресурсы: {ex.Message}", "VoicePilot",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    [RelayCommand]
    private void Cancel() => CloseRequested?.Invoke(this, false);

    [RelayCommand]
    private void Reset()
    {
        Load();
    }

    private void Load()
    {
        Resources.Clear();

        foreach (var resource in _resourceRegistry.GetAll())
        {
            Resources.Add(ResourceEditorItem.FromDescriptor(resource));
        }

        SelectedResource = Resources.FirstOrDefault();
    }

    public class ResourceEditorItem : ObservableObject
    {
        private string _key = string.Empty;
        private string _path = string.Empty;
        private string _type = "Folder";
        private string? _description;

        public string Key
        {
            get => _key;
            set => SetProperty(ref _key, value);
        }

        public string Path
        {
            get => _path;
            set => SetProperty(ref _path, value);
        }

        public string Type
        {
            get => _type;
            set => SetProperty(ref _type, value);
        }

        public string? Description
        {
            get => _description;
            set => SetProperty(ref _description, value);
        }

        public ResourceDescriptor ToDescriptor() => new()
        {
            Key = Key,
            Path = Path,
            Type = Type,
            Description = Description
        };

        public static ResourceEditorItem FromDescriptor(ResourceDescriptor descriptor) => new()
        {
            Key = descriptor.Key,
            Path = descriptor.Path,
            Type = descriptor.Type,
            Description = descriptor.Description
        };
    }
}
