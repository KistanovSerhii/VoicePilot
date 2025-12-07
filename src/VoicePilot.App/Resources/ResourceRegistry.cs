using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Encodings.Web;
using System.Text.Json;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using VoicePilot.App.Configuration;

namespace VoicePilot.App.Resources;

public class ResourceRegistry
{
    private readonly string _resourceFile;
    private readonly ILogger<ResourceRegistry> _logger;
    private readonly object _sync = new();

    private Dictionary<string, ResourceDescriptor> _resources =
        new(StringComparer.OrdinalIgnoreCase);

    public ResourceRegistry(IOptions<PathOptions> options, ILogger<ResourceRegistry> logger)
    {
        _resourceFile = options.Value.ResourcesFile;
        _logger = logger;

        LoadResources();
    }

    public IReadOnlyCollection<ResourceDescriptor> GetAll()
    {
        lock (_sync)
        {
            return _resources.Values.Select(r => r).ToList();
        }
    }

    public bool TryGet(string key, out ResourceDescriptor descriptor)
    {
        lock (_sync)
        {
            return _resources.TryGetValue(key, out descriptor!);
        }
    }

    public void SaveResources(IEnumerable<ResourceDescriptor> resources)
    {
        if (resources is null)
        {
            throw new ArgumentNullException(nameof(resources));
        }

        var normalised = Normalise(resources);

        lock (_sync)
        {
            _resources = normalised.ToDictionary(r => r.Key, r => r, StringComparer.OrdinalIgnoreCase);
        }

        var payload = new ResourceFile
        {
            Resources = normalised
                .Select(r => new ResourceFileItem
                {
                    Key = r.Key,
                    Path = r.Path,
                    Type = r.Type,
                    Description = r.Description
                })
                .ToList()
        };

        Directory.CreateDirectory(Path.GetDirectoryName(_resourceFile)!);

        using var stream = File.Create(_resourceFile);
        JsonSerializer.Serialize(stream, payload, new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        });
    }

    private void LoadResources()
    {
        try
        {
            if (!File.Exists(_resourceFile))
            {
                WriteDefaultResources();
            }

            using var stream = File.OpenRead(_resourceFile);
            var payload = JsonSerializer.Deserialize<ResourceFile>(stream);

            if (payload is null)
            {
                _logger.LogWarning("Resources file {File} is empty.", _resourceFile);
                return;
            }

            lock (_sync)
            {
                _resources = new Dictionary<string, ResourceDescriptor>(StringComparer.OrdinalIgnoreCase);

                foreach (var resource in payload.Resources)
                {
                    if (string.IsNullOrWhiteSpace(resource.Key) || string.IsNullOrWhiteSpace(resource.Path))
                    {
                        continue;
                    }

                    _resources[resource.Key] = new ResourceDescriptor
                    {
                        Key = resource.Key,
                        Path = resource.Path,
                        Type = resource.Type ?? "Generic",
                        Description = resource.Description
                    };
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to load resources from {File}.", _resourceFile);
        }
    }

    private void WriteDefaultResources()
    {
        var payload = new ResourceFile
        {
            Resources =
            [
                new ResourceFileItem
                {
                    Key = "фильмы",
                    Path = "%USERPROFILE%\\Videos",
                    Type = "Folder",
                    Description = "Папка пользователя с видео"
                },
                new ResourceFileItem
                {
                    Key = "музыка",
                    Path = "%USERPROFILE%\\Music",
                    Type = "Folder",
                    Description = "Папка пользователя с музыкой"
                }
            ]
        };

        Directory.CreateDirectory(Path.GetDirectoryName(_resourceFile)!);

        using var stream = File.Create(_resourceFile);
        JsonSerializer.Serialize(stream, payload, new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        });
    }

    private static List<ResourceDescriptor> Normalise(IEnumerable<ResourceDescriptor> items)
    {
        return items
            .Where(item => !string.IsNullOrWhiteSpace(item.Key) && !string.IsNullOrWhiteSpace(item.Path))
            .Select(item => new ResourceDescriptor
            {
                Key = item.Key.Trim(),
                Path = item.Path.Trim(),
                Type = string.IsNullOrWhiteSpace(item.Type) ? "Generic" : item.Type.Trim(),
                Description = string.IsNullOrWhiteSpace(item.Description) ? null : item.Description.Trim()
            })
            .GroupBy(item => item.Key, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .ToList();
    }

    private class ResourceFile
    {
        public List<ResourceFileItem> Resources { get; set; } = new();
    }

    private class ResourceFileItem
    {
        public string Key { get; set; } = string.Empty;

        public string Path { get; set; } = string.Empty;

        public string? Type { get; set; }

        public string? Description { get; set; }
    }
}
