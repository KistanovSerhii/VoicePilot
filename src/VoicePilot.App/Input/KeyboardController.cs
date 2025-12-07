using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Extensions.Logging;
using WindowsInput;
using WindowsInput.Events;

namespace VoicePilot.App.Input;

public class KeyboardController
{
    private readonly ILogger<KeyboardController> _logger;

    public KeyboardController(ILogger<KeyboardController> logger)
    {
        _logger = logger;
    }

    public void SendShortcut(IEnumerable<string> keyTokens)
    {
        var keys = keyTokens
            .Select(ParseKey)
            .Where(key => key is not null)
            .Cast<KeyCode>()
            .ToArray();

        if (keys.Length == 0)
        {
            return;
        }

        Simulate.Events()
            .ClickChord(keys)
            .Invoke()
            .GetAwaiter()
            .GetResult();
    }

    public void TypeText(string text)
    {
        Simulate.Events()
            .Click(text)
            .Invoke()
            .GetAwaiter()
            .GetResult();
    }

    private KeyCode? ParseKey(string token)
    {
        if (string.IsNullOrWhiteSpace(token))
        {
            return null;
        }

        var normalized = token.Trim().Replace(" ", string.Empty).ToLowerInvariant();

        return normalized switch
        {
            "ctrl" or "control" => KeyCode.Control,
            "shift" => KeyCode.Shift,
            "alt" => KeyCode.Alt,
            "lalt" => KeyCode.LMenu,
            "ralt" => KeyCode.RMenu,
            "win" or "windows" or "lwin" => KeyCode.LWin,
            "rwin" => KeyCode.RWin,
            "enter" or "return" => KeyCode.Return,
            "esc" or "escape" => KeyCode.Escape,
            "space" or "spacebar" => KeyCode.Space,
            "tab" => KeyCode.Tab,
            "up" => KeyCode.Up,
            "down" => KeyCode.Down,
            "left" => KeyCode.Left,
            "right" => KeyCode.Right,
            "home" => KeyCode.Home,
            "end" => KeyCode.End,
            "pageup" or "pgup" => KeyCode.Prior,
            "pagedown" or "pgdn" => KeyCode.Next,
            "delete" or "del" => KeyCode.Delete,
            "backspace" or "bs" => KeyCode.Backspace,
            "f1" or "f2" or "f3" or "f4" or "f5" or "f6" or "f7" or "f8" or "f9" or "f10" or "f11" or "f12" => ParseFunctionKey(normalized),
            _ => ParseCharacterKey(normalized)
        };
    }

    private KeyCode? ParseFunctionKey(string key)
    {
        if (int.TryParse(key.AsSpan(1), out var number) && number is >= 1 and <= 24)
        {
            var name = $"F{number}";
            if (Enum.TryParse(name, out KeyCode result))
            {
                return result;
            }
        }

        return null;
    }

    private KeyCode? ParseCharacterKey(string token)
    {
        if (token.Length == 1)
        {
            var ch = token[0];
            if (char.IsLetter(ch))
            {
                var name = char.ToUpperInvariant(ch).ToString();
                if (Enum.TryParse(name, out KeyCode letterCode))
                {
                    return letterCode;
                }
            }

            if (char.IsDigit(ch))
            {
                var name = $"D{ch}";
                if (Enum.TryParse(name, out KeyCode digitCode))
                {
                    return digitCode;
                }
            }
        }

        if (Enum.TryParse(token, true, out KeyCode code))
        {
            return code;
        }

        _logger.LogWarning("Unsupported key token '{Token}'.", token);
        return null;
    }
}
