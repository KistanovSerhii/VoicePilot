using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using Microsoft.Extensions.Logging;

namespace VoicePilot.App.Modules;

public class ModuleManifestParser
{
    private readonly ILogger<ModuleManifestParser> _logger;

    public ModuleManifestParser(ILogger<ModuleManifestParser> logger)
    {
        _logger = logger;
    }

    public CommandModule Parse(string manifestPath)
    {
        if (!File.Exists(manifestPath))
        {
            throw new FileNotFoundException("Manifest file not found.", manifestPath);
        }

        var extension = Path.GetExtension(manifestPath).ToLowerInvariant();

        return extension switch
        {
            ".json" => ParseJson(manifestPath),
            ".xml" => ParseXml(manifestPath),
            _ => throw new NotSupportedException($"Unsupported manifest format '{extension}'.")
        };
    }

    private CommandModule ParseJson(string manifestPath)
    {
        using var stream = File.OpenRead(manifestPath);
        var root = JsonNode.Parse(stream)?.AsObject()
                   ?? throw new InvalidDataException("Manifest JSON root is invalid.");

        var module = new CommandModule
        {
            Id = ReadString(root, "id", Path.GetFileName(Path.GetDirectoryName(manifestPath)!)),
            Name = ReadString(root, "name", "Unnamed module"),
            Version = ReadString(root, "version", "1.0.0"),
            Culture = ReadString(root, "culture", "ru-RU"),
            Author = ReadString(root, "author", "Unknown author"),
            Description = ReadString(root, "description", string.Empty),
            BaseDirectory = Path.GetDirectoryName(manifestPath) ?? AppContext.BaseDirectory,
            Commands = ParseJsonCommands(root),
            Resources = ParseJsonResources(root)
        };

        return module;
    }

    private static IReadOnlyList<VoiceCommand> ParseJsonCommands(JsonObject manifest)
    {
        if (!manifest.TryGetPropertyValue("commands", out var node) || node is not JsonArray array)
        {
            return Array.Empty<VoiceCommand>();
        }

        var commands = new List<VoiceCommand>();

        foreach (var item in array.OfType<JsonObject>())
        {
            commands.Add(ParseJsonCommand(item));
        }

        return commands;
    }

    private static VoiceCommand ParseJsonCommand(JsonObject obj)
    {
        var actions = new List<CommandAction>();

        if (obj.TryGetPropertyValue("actions", out var actionsNode) && actionsNode is JsonArray actionArray)
        {
            foreach (var actionNode in actionArray.OfType<JsonObject>())
            {
                actions.Add(ParseJsonAction(actionNode));
            }
        }

        var phrases = new List<string>();
        if (obj.TryGetPropertyValue("phrases", out var phrasesNode) && phrasesNode is JsonArray phraseArray)
        {
            foreach (var phrase in phraseArray)
            {
                if (phrase is JsonValue value && value.TryGetValue<string>(out var text) && !string.IsNullOrWhiteSpace(text))
                {
                    phrases.Add(text.Trim());
                }
            }
        }

        return new VoiceCommand
        {
            Id = ReadString(obj, "id", Guid.NewGuid().ToString("N")),
            Name = ReadString(obj, "name", string.Empty),
            Description = ReadString(obj, "description", string.Empty),
            Phrases = phrases,
            Actions = actions,
            CustomThreshold = ReadNullableDouble(obj, "threshold"),
            KeepAssistantActive = ReadBool(obj, "keepActive")
        };
    }

    private static CommandAction ParseJsonAction(JsonObject obj)
    {
        var parameters = obj.TryGetPropertyValue("parameters", out var parameterNode) && parameterNode is JsonObject parameterObject
            ? CloneJsonObject(parameterObject)
            : new JsonObject();

        var nested = new List<CommandAction>();

        if (obj.TryGetPropertyValue("actions", out var nestedNode) && nestedNode is JsonArray nestedArray)
        {
            foreach (var nestedObject in nestedArray.OfType<JsonObject>())
            {
                nested.Add(ParseJsonAction(nestedObject));
            }
        }

        return new CommandAction(ReadString(obj, "type", string.Empty), parameters, nested);
    }

    private static IReadOnlyDictionary<string, ModuleResource> ParseJsonResources(JsonObject manifest)
    {
        if (!manifest.TryGetPropertyValue("resources", out var node) || node is not JsonArray array)
        {
            return new Dictionary<string, ModuleResource>(StringComparer.OrdinalIgnoreCase);
        }

        var resources = new Dictionary<string, ModuleResource>(StringComparer.OrdinalIgnoreCase);

        foreach (var item in array.OfType<JsonObject>())
        {
            var resource = new ModuleResource
            {
                Key = ReadString(item, "key", ReadString(item, "name", string.Empty)),
                Path = ReadString(item, "path", string.Empty),
                Type = ReadString(item, "type", "Generic")
            };

            if (!string.IsNullOrWhiteSpace(resource.Key))
            {
                resources[resource.Key] = resource;
            }
        }

        return resources;
    }

    private CommandModule ParseXml(string manifestPath)
    {
        var document = XDocument.Load(manifestPath);
        var root = document.Root ?? throw new InvalidDataException("Manifest XML root is missing.");

        static string ReadElement(XElement element, string name, string defaultValue = "")
        {
            return element.Element(name)?.Value?.Trim() ?? defaultValue;
        }

        var module = new CommandModule
        {
            Id = root.Attribute("id")?.Value ?? ReadElement(root, "id", Path.GetFileName(Path.GetDirectoryName(manifestPath)!)),
            Name = root.Attribute("name")?.Value ?? ReadElement(root, "name", "Unnamed module"),
            Version = root.Attribute("version")?.Value ?? ReadElement(root, "version", "1.0.0"),
            Culture = ReadElement(root, "culture", "ru-RU"),
            Author = ReadElement(root, "author", "Unknown author"),
            Description = ReadElement(root, "description", string.Empty),
            BaseDirectory = Path.GetDirectoryName(manifestPath) ?? AppContext.BaseDirectory,
            Commands = ParseXmlCommands(root),
            Resources = ParseXmlResources(root)
        };

        return module;
    }

    private static IReadOnlyList<VoiceCommand> ParseXmlCommands(XElement root)
    {
        var commandsElement = root.Element("commands");
        if (commandsElement is null)
        {
            return Array.Empty<VoiceCommand>();
        }

        var commands = new List<VoiceCommand>();

        foreach (var commandElement in commandsElement.Elements("command"))
        {
            var phrases = commandElement.Element("phrases")?.Elements("phrase")
                .Select(p => p.Value.Trim())
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .ToList() ?? new List<string>();

            var actions = new List<CommandAction>();
            var actionsElement = commandElement.Element("actions");
            if (actionsElement is not null)
            {
                foreach (var actionElement in actionsElement.Elements("action"))
                {
                    actions.Add(ParseXmlAction(actionElement));
                }
            }

            commands.Add(new VoiceCommand
            {
                Id = commandElement.Attribute("id")?.Value ?? Guid.NewGuid().ToString("N"),
                Name = commandElement.Attribute("name")?.Value ?? (commandElement.Element("name")?.Value ?? string.Empty),
                Description = commandElement.Element("description")?.Value ?? string.Empty,
                Phrases = phrases,
                Actions = actions,
                CustomThreshold = ReadNullableDouble(commandElement, "threshold"),
                KeepAssistantActive = ReadBool(commandElement, "keepActive")
            });
        }

        return commands;
    }

    private static CommandAction ParseXmlAction(XElement element)
    {
        var parameters = new JsonObject();

        foreach (var attribute in element.Attributes())
        {
            if (!attribute.IsNamespaceDeclaration && !string.Equals(attribute.Name.LocalName, "type", StringComparison.OrdinalIgnoreCase))
            {
                parameters[attribute.Name.LocalName] = attribute.Value;
            }
        }

        foreach (var child in element.Elements())
        {
            if (string.Equals(child.Name.LocalName, "actions", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (!child.HasElements)
            {
                parameters[child.Name.LocalName] = child.Value;
            }
        }

        var nested = new List<CommandAction>();
        var nestedElement = element.Element("actions");
        if (nestedElement is not null)
        {
            foreach (var child in nestedElement.Elements("action"))
            {
                nested.Add(ParseXmlAction(child));
            }
        }

        var type = element.Attribute("type")?.Value ?? element.Element("type")?.Value ?? string.Empty;
        return new CommandAction(type, parameters, nested);
    }

    private static IReadOnlyDictionary<string, ModuleResource> ParseXmlResources(XElement root)
    {
        var resourcesElement = root.Element("resources");
        if (resourcesElement is null)
        {
            return new Dictionary<string, ModuleResource>(StringComparer.OrdinalIgnoreCase);
        }

        var resources = new Dictionary<string, ModuleResource>(StringComparer.OrdinalIgnoreCase);

        foreach (var resourceElement in resourcesElement.Elements("resource"))
        {
            var key = resourceElement.Attribute("key")?.Value
                      ?? resourceElement.Element("key")?.Value
                      ?? resourceElement.Attribute("name")?.Value
                      ?? resourceElement.Element("name")?.Value
                      ?? string.Empty;

            if (string.IsNullOrWhiteSpace(key))
            {
                continue;
            }

            resources[key] = new ModuleResource
            {
                Key = key,
                Path = resourceElement.Attribute("path")?.Value ?? resourceElement.Element("path")?.Value ?? string.Empty,
                Type = resourceElement.Attribute("type")?.Value ?? resourceElement.Element("type")?.Value ?? "Generic"
            };
        }

        return resources;
    }

    private static JsonObject CloneJsonObject(JsonObject source)
    {
        var clone = new JsonObject();
        foreach (var property in source)
        {
            clone[property.Key] = property.Value?.DeepClone();
        }

        return clone;
    }

    private static string ReadString(JsonObject obj, string property, string defaultValue)
    {
        if (obj.TryGetPropertyValue(property, out var node) &&
            node is JsonValue value &&
            value.TryGetValue<string>(out var result) &&
            !string.IsNullOrWhiteSpace(result))
        {
            return result.Trim();
        }

        return defaultValue;
    }

    private static double? ReadNullableDouble(JsonObject obj, string property)
    {
        if (obj.TryGetPropertyValue(property, out var node) &&
            node is JsonValue value &&
            value.TryGetValue<double>(out var result))
        {
            return result;
        }

        if (obj.TryGetPropertyValue(property, out node) &&
            node is JsonValue altValue &&
            altValue.TryGetValue<string>(out var text) &&
            double.TryParse(text, out var parsed))
        {
            return parsed;
        }

        return null;
    }

    private static double? ReadNullableDouble(XElement element, string attributeName)
    {
        if (double.TryParse(element.Attribute(attributeName)?.Value, out var fromAttribute))
        {
            return fromAttribute;
        }

        if (double.TryParse(element.Element(attributeName)?.Value, out var fromElement))
        {
            return fromElement;
        }

        return null;
    }

    private static bool ReadBool(JsonObject obj, string property)
    {
        if (obj.TryGetPropertyValue(property, out var node))
        {
            if (node is JsonValue value)
            {
                if (value.TryGetValue<bool>(out var boolean))
                {
                    return boolean;
                }

                if (value.TryGetValue<string>(out var text) &&
                    bool.TryParse(text, out var parsed))
                {
                    return parsed;
                }
            }
        }

        return false;
    }

    private static bool ReadBool(XElement element, string attributeName)
    {
        if (bool.TryParse(element.Attribute(attributeName)?.Value, out var fromAttribute))
        {
            return fromAttribute;
        }

        if (bool.TryParse(element.Element(attributeName)?.Value, out var fromElement))
        {
            return fromElement;
        }

        return false;
    }
}
