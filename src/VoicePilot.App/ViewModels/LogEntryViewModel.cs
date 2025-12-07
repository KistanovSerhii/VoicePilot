using System;
using Microsoft.Extensions.Logging;

namespace VoicePilot.App.ViewModels;

public class LogEntryViewModel
{
    public LogEntryViewModel(LogLevel level, string message)
    {
        Timestamp = DateTimeOffset.Now;
        Level = level;
        Message = message;
    }

    public DateTimeOffset Timestamp { get; }

    public LogLevel Level { get; }

    public string Message { get; }
}
