using VoicePilot.App.Modules;

namespace VoicePilot.App.ViewModels;

public class ModuleViewModel
{
    public ModuleViewModel(CommandModule module)
    {
        Id = module.Id;
        Name = module.Name;
        Version = module.Version;
        Author = module.Author;
        Description = module.Description;
        Culture = module.Culture;
        CommandsCount = module.Commands.Count;
    }

    public string Id { get; }

    public string Name { get; }

    public string Version { get; }

    public string Author { get; }

    public string Description { get; }

    public string Culture { get; }

    public int CommandsCount { get; }
}
