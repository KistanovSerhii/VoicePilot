using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Encodings.Web;
using System.Text.Json;
using FuzzySharp;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using VoicePilot.App.Configuration;

namespace VoicePilot.App.Services;

public class SystemCommandService
{
    private readonly string _configPath;
    private readonly ILogger<SystemCommandService> _logger;

    private SystemCommandManifest _manifest = new();
    private readonly object _sync = new();

    private const double MatchThreshold = 0.82;

    public SystemCommandService(IOptions<PathOptions> options, ILogger<SystemCommandService> logger)
    {
        _configPath = options.Value.SystemCommandsFile;
        _logger = logger;

        LoadManifest();
    }

    public IReadOnlyCollection<string> ActivationKeywords => _manifest.Activation;

    public IReadOnlyCollection<string> SilenceKeywords => _manifest.Silence;

    public IReadOnlyCollection<string> ExitKeywords => _manifest.Exit;

    public SystemCommandManifest GetManifestSnapshot()
    {
        lock (_sync)
        {
            return CloneManifest(_manifest);
        }
    }

    public void SaveManifest(SystemCommandManifest manifest)
    {
        if (manifest is null)
        {
            throw new ArgumentNullException(nameof(manifest));
        }

        var normalised = NormaliseManifest(manifest);

        lock (_sync)
        {
            _manifest = normalised;
        }

        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(_configPath)!);
            using var stream = File.Create(_configPath);
            JsonSerializer.Serialize(stream, _manifest, new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to save system command manifest to {File}.", _configPath);
            throw;
        }
    }

    public SystemCommandType Match(string recognisedText, out double score)
    {
        var normalised = Normalise(recognisedText);

        SystemCommandType bestType = SystemCommandType.None;
        var bestScore = 0.0;

        Evaluate(normalised, _manifest.Activation, SystemCommandType.Activation, ref bestType, ref bestScore);
        Evaluate(normalised, _manifest.Silence, SystemCommandType.Silence, ref bestType, ref bestScore);
        Evaluate(normalised, _manifest.Exit, SystemCommandType.Exit, ref bestType, ref bestScore);

        score = bestScore;
        return bestScore >= MatchThreshold ? bestType : SystemCommandType.None;
    }

    private void Evaluate(
        string recognised,
        IEnumerable<string> phrases,
        SystemCommandType type,
        ref SystemCommandType bestType,
        ref double bestScore)
    {
        foreach (var phrase in phrases)
        {
            if (string.IsNullOrWhiteSpace(phrase))
            {
                continue;
            }

            var candidate = Normalise(phrase);
            double score;

            if (string.Equals(candidate, recognised, StringComparison.OrdinalIgnoreCase))
            {
                score = 1.0;
            }
            else
            {
                score = Fuzz.WeightedRatio(recognised, candidate) / 100.0;
            }

            if (score > bestScore)
            {
                bestScore = score;
                bestType = type;
            }
        }
    }

    private void LoadManifest()
    {
        try
        {
            if (!File.Exists(_configPath))
            {
                WriteDefaultManifest();
            }

            using var stream = File.OpenRead(_configPath);
            var manifest = JsonSerializer.Deserialize<SystemCommandManifest>(stream);
            if (manifest is not null)
            {
                lock (_sync)
                {
                    _manifest = NormaliseManifest(manifest);
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to load system command manifest from {File}.", _configPath);
        }
    }

    private void WriteDefaultManifest()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(_configPath)!);

        var manifest = new SystemCommandManifest
        {
            Activation = { "командир", "слушай", "ассистент" },
            Silence = { "тишина", "не слушай" },
            Exit = { "выход", "заверши работу" }
        };

        using var stream = File.Create(_configPath);
        JsonSerializer.Serialize(stream, manifest, new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        });

        lock (_sync)
        {
            _manifest = manifest;
        }
    }

    private static string Normalise(string value)
    {
        return value.Trim().ToLowerInvariant();
    }

    private static SystemCommandManifest CloneManifest(SystemCommandManifest source) =>
        new()
        {
            Activation = source.Activation.ToList(),
            Silence = source.Silence.ToList(),
            Exit = source.Exit.ToList()
        };

    private static SystemCommandManifest NormaliseManifest(SystemCommandManifest manifest) =>
        new()
        {
            Activation = NormaliseList(manifest.Activation),
            Silence = NormaliseList(manifest.Silence),
            Exit = NormaliseList(manifest.Exit)
        };

    private static List<string> NormaliseList(IEnumerable<string> source) =>
        source
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .Select(s => s.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
}






