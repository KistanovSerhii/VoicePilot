using VoicePilot.App.ViewModels;

namespace VoicePilot.App.Windows;

public class CommandBuilderContext
{
    private CommandBuilderContext()
    {
    }

    public bool IsEditMode => SourceItem is not null;

    public CommandListItemViewModel? SourceItem { get; private set; }

    public static CommandBuilderContext CreateNew() => new CommandBuilderContext();

    public static CommandBuilderContext ForExisting(CommandListItemViewModel item)
        => new CommandBuilderContext { SourceItem = item };
}
