using System;
using Microsoft.Extensions.Logging;

namespace VoicePilot.App.Messaging;

public record RecognitionLogMessage(LogLevel Level, string Message, Exception? Exception = null);
