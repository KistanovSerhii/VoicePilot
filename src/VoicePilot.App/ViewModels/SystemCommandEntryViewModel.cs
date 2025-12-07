using CommunityToolkit.Mvvm.ComponentModel;
using VoicePilot.App.Services;

namespace VoicePilot.App.ViewModels;

public partial class SystemCommandEntryViewModel : ObservableObject
{
    public SystemCommandEntryViewModel(string title, string keywords, SystemCommandType type)
    {
        Title = title;
        Keywords = keywords;
        Type = type;
    }

    public string Title { get; }

    public string Keywords { get; }

    public SystemCommandType Type { get; }

    [ObservableProperty]
    private bool _isCurrent;
}
