using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using CommunityToolkit.Mvvm.Messaging;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using VoicePilot.App.Configuration;
using VoicePilot.App.Messaging;

namespace VoicePilot.App.Modules;

public class ModuleManager
{
    private readonly ModuleManifestParser _parser;
    private readonly IOptions<PathOptions> _pathOptions;
    private readonly IMessenger _messenger;
    private readonly ILogger<ModuleManager> _logger;

    private readonly List<CommandModule> _modules = new();
    private readonly SemaphoreSlim _semaphore = new(1, 1);

    public ModuleManager(
        ModuleManifestParser parser,
        IOptions<PathOptions> pathOptions,
        IMessenger messenger,
        ILogger<ModuleManager> logger)
    {
        _parser = parser;
        _pathOptions = pathOptions;
        _messenger = messenger;
        _logger = logger;
    }

    public IReadOnlyList<CommandModule> Modules
    {
        get
        {
            lock (_modules)
            {
                return _modules.ToList();
            }
        }
    }

    public async Task InitialiseAsync(CancellationToken cancellationToken)
    {
        await _semaphore.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            _modules.Clear();

            var modulesDirectory = _pathOptions.Value.ModulesDirectory;
            if (!Directory.Exists(modulesDirectory))
            {
                Directory.CreateDirectory(modulesDirectory);
            }

            TrySeedModules(modulesDirectory);

            foreach (var directory in Directory.EnumerateDirectories(modulesDirectory))
            {
                cancellationToken.ThrowIfCancellationRequested();
                TryLoadModule(directory);
            }

            BroadcastCatalog();
        }
        finally
        {
            _semaphore.Release();
        }
    }

    public async Task<CommandModule?> ImportModuleAsync(string archivePath, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(archivePath))
        {
            throw new ArgumentException("Archive path must be provided.", nameof(archivePath));
        }

        if (!File.Exists(archivePath))
        {
            throw new FileNotFoundException("Module archive was not found.", archivePath);
        }

        await _semaphore.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            var stagingRoot = _pathOptions.Value.ModuleStagingDirectory;
            Directory.CreateDirectory(stagingRoot);

            var stagingDirectory = Path.Combine(stagingRoot, Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(stagingDirectory);

            ZipFile.ExtractToDirectory(archivePath, stagingDirectory, overwriteFiles: true);

            var manifestPath = FindManifest(stagingDirectory);
            if (manifestPath is null)
            {
                throw new InvalidDataException("The module archive does not contain a manifest.json or manifest.xml file.");
            }

            var manifest = _parser.Parse(manifestPath);
            var targetDirectory = BuildTargetDirectory(manifest);

            if (Directory.Exists(targetDirectory))
            {
                Directory.Delete(targetDirectory, recursive: true);
            }

            Directory.CreateDirectory(Path.GetDirectoryName(targetDirectory)!);
            var moduleSourceDirectory = Path.GetDirectoryName(manifestPath)!;
            Directory.Move(moduleSourceDirectory, targetDirectory);
            TryDeleteQuietly(stagingDirectory);

            TryLoadModule(targetDirectory);
            BroadcastCatalog();

            return _modules.FirstOrDefault(m => string.Equals(m.Id, manifest.Id, StringComparison.OrdinalIgnoreCase));
        }
        finally
        {
            _semaphore.Release();
        }
    }

    public CommandModule? GetModule(string moduleId)
    {
        lock (_modules)
        {
            return _modules.FirstOrDefault(m => string.Equals(m.Id, moduleId, StringComparison.OrdinalIgnoreCase));
        }
    }

    private void TryLoadModule(string directory)
    {
        var manifestPath = FindManifest(directory);

        if (manifestPath is null)
        {
            _logger.LogWarning("Skipping module at {Directory}: manifest not found.", directory);
            return;
        }

        try
        {
            var module = _parser.Parse(manifestPath);

            lock (_modules)
            {
                var existingIndex = _modules.FindIndex(m => string.Equals(m.Id, module.Id, StringComparison.OrdinalIgnoreCase));
                if (existingIndex >= 0)
                {
                    _modules[existingIndex] = module;
                }
                else
                {
                    _modules.Add(module);
                }
            }

            _logger.LogInformation("Loaded command module {ModuleName} ({ModuleId}) from {Directory}.", module.Name, module.Id, directory);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to load module from {Directory}.", directory);
        }
    }

    private void BroadcastCatalog()
    {
        var snapshot = Modules;
        _messenger.Send(new ModuleCatalogChangedMessage(snapshot));
    }

    private void TrySeedModules(string modulesDirectory)
    {
        if (Directory.EnumerateDirectories(modulesDirectory).Any())
        {
            return;
        }

        var candidate = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "Modules"));
        if (!Directory.Exists(candidate))
        {
            return;
        }

        foreach (var sourceDir in Directory.EnumerateDirectories(candidate))
        {
            var targetDir = Path.Combine(modulesDirectory, Path.GetFileName(sourceDir));
            if (Directory.Exists(targetDir))
            {
                continue;
            }

            try
            {
                CopyDirectory(sourceDir, targetDir);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to seed module from {Source}", sourceDir);
            }
        }
    }

    private static void CopyDirectory(string source, string destination)
    {
        Directory.CreateDirectory(destination);

        foreach (var file in Directory.EnumerateFiles(source))
        {
            var targetFile = Path.Combine(destination, Path.GetFileName(file));
            File.Copy(file, targetFile, overwrite: true);
        }

        foreach (var directory in Directory.EnumerateDirectories(source))
        {
            var targetSubdirectory = Path.Combine(destination, Path.GetFileName(directory));
            CopyDirectory(directory, targetSubdirectory);
        }
    }

    private static string? FindManifest(string directory)
    {
        var jsonManifest = Path.Combine(directory, "manifest.json");
        if (File.Exists(jsonManifest))
        {
            return jsonManifest;
        }

        var xmlManifest = Path.Combine(directory, "manifest.xml");
        if (File.Exists(xmlManifest))
        {
            return xmlManifest;
        }

        return null;
    }

    private string BuildTargetDirectory(CommandModule module)
    {
        var baseDirectory = _pathOptions.Value.ModulesDirectory;
        var folderName = $"{Sanitise(module.Id)}-{Sanitise(module.Version)}";
        return Path.Combine(baseDirectory, folderName);
    }

    private static string Sanitise(string value)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var chars = value.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray();
        return new string(chars);
    }

    private static void TryDeleteQuietly(string directory)
    {
        try
        {
            if (Directory.Exists(directory))
            {
                Directory.Delete(directory, recursive: true);
            }
        }
        catch
        {
            // ignore cleanup errors
        }
    }
}
