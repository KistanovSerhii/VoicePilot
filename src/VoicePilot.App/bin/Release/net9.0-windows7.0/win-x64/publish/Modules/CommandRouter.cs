using System;
using System.Linq;
using FuzzySharp;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using VoicePilot.App.Configuration;

namespace VoicePilot.App.Modules;

public class CommandRouter
{
    private readonly ModuleManager _moduleManager;
    private readonly IOptions<SpeechOptions> _speechOptions;
    private readonly ILogger<CommandRouter> _logger;

    public CommandRouter(
        ModuleManager moduleManager,
        IOptions<SpeechOptions> speechOptions,
        ILogger<CommandRouter> logger)
    {
        _moduleManager = moduleManager;
        _speechOptions = speechOptions;
        _logger = logger;
    }

    public bool TryMatch(string recognisedText, out CommandMatch match)
    {
        match = null!;

        if (string.IsNullOrWhiteSpace(recognisedText))
        {
            return false;
        }

        var normalizedInput = Normalise(recognisedText);
        var options = _speechOptions.Value;

        CommandMatch? bestMatch = null;
        var bestScore = 0.0;

        foreach (var module in _moduleManager.Modules)
        {
            foreach (var command in module.Commands)
            {
                foreach (var phrase in command.Phrases)
                {
                    var normalizedPhrase = Normalise(phrase);
                    double score;

                    if (string.Equals(normalizedPhrase, normalizedInput, StringComparison.OrdinalIgnoreCase))
                    {
                        score = 1.0;
                    }
                    else
                    {
                        score = Fuzz.WeightedRatio(normalizedInput, normalizedPhrase) / 100.0;
                    }

                    if (score > bestScore)
                    {
                        bestScore = score;
                        bestMatch = new CommandMatch(module, command, phrase, score);
                    }
                }
            }
        }

        if (bestMatch is null)
        {
            return false;
        }

        var threshold = bestMatch.Command.CustomThreshold ?? options.CommandMatchThreshold;

        if (bestScore >= threshold)
        {
            match = bestMatch;
            return true;
        }

        if (bestScore >= options.CommandMatchFallbackThreshold)
        {
            _logger.LogDebug("Rejected command '{Text}' matched phrase '{Phrase}' with score {Score:0.00}, below threshold {Threshold:0.00}.",
                recognisedText, bestMatch.MatchedPhrase, bestScore, threshold);
        }

        return false;
    }

    private static string Normalise(string phrase)
    {
        var trimmed = phrase.Trim();
        return trimmed.ToLowerInvariant();
    }
}
