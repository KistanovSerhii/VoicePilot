using System;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Extensions.Logging;
using VoicePilot.App.Input;
using VoicePilot.App.Modules;
using VoicePilot.App.Resources;
using VoicePilot.App.Windows;

namespace VoicePilot.App.Services;

public class ActionExecutor
{
    private readonly MouseController _mouseController;
    private readonly KeyboardController _keyboardController;
    private readonly ResourceRegistry _resourceRegistry;
    private readonly SpeechSynthesisService _speechSynthesis;
    private readonly ILogger<ActionExecutor> _logger;
    private readonly ConcurrentDictionary<string, List<Process>> _processes = new();
    private readonly ConcurrentDictionary<string, List<string>> _explorerFolders = new();
    private readonly ConcurrentDictionary<string, List<WebViewWindow>> _webViewWindows = new();
    private readonly ILoggerFactory _loggerFactory;

    public ActionExecutor(
        MouseController mouseController,
        KeyboardController keyboardController,
        ResourceRegistry resourceRegistry,
        SpeechSynthesisService speechSynthesis,
        ILogger<ActionExecutor> logger,
        ILoggerFactory loggerFactory)
    {
        _mouseController = mouseController;
        _keyboardController = keyboardController;
        _resourceRegistry = resourceRegistry;
        _speechSynthesis = speechSynthesis;
        _logger = logger;
        _loggerFactory = loggerFactory;
    }

    public async Task ExecuteAsync(CommandModule module, VoiceCommand command, CancellationToken cancellationToken)
    {
        foreach (var action in command.Actions)
        {
            await ExecuteActionAsync(module, command, action, cancellationToken).ConfigureAwait(false);
        }
    }

    private async Task ExecuteActionAsync(CommandModule module, VoiceCommand command, CommandAction action, CancellationToken cancellationToken)
    {
        var type = action.Type.ToLowerInvariant();

        switch (type)
        {
            case "runprocess":
            case "run":
                RunProcess(module, command, action.Parameters);
                break;
            case "shell":
                RunShellCommand(action.Parameters);
                break;
            case "script":
                RunScript(module, action.Parameters);
                break;
            case "stopprocess":
                StopProcess(module, command, action.Parameters);
                break;
            case "stopallprocesses":
            case "stopall":
                StopAllProcesses(module, command, action.Parameters);
                break;
            case "sendkeys":
            case "keypress":
            case "shortcut":
                SendKeys(action.Parameters);
                break;
            case "typetext":
            case "type":
            case "text":
                TypeText(action.Parameters);
                break;
            case "mousemove":
                HandleMouseMove(action.Parameters, cancellationToken);
                break;
            case "mousestop":
            case "stopmouse":
                _mouseController.StopContinuousMove();
                break;
            case "mouseclick":
                HandleMouseClick(action.Parameters);
                break;
            case "openurl":
                OpenUrl(action.Parameters);
                break;
            case "openwebview":
                OpenWebView(module, command, action.Parameters);
                break;
            case "openresource":
                OpenResource(module, command, action.Parameters);
                break;
            case "speak":
                await SpeakAsync(action.Parameters, cancellationToken).ConfigureAwait(false);
                break;
            case "pause":
            case "delay":
                await PauseAsync(action.Parameters, cancellationToken).ConfigureAwait(false);
                break;
            default:
                _logger.LogWarning("Unknown action type '{ActionType}'.", action.Type);
                break;
        }

        if (action.Children.Count > 0)
        {
            foreach (var child in action.Children)
            {
                await ExecuteActionAsync(module, command, child, cancellationToken).ConfigureAwait(false);
            }
        }
    }

    private void RunProcess(CommandModule module, VoiceCommand command, JsonObject parameters)
    {
        var path = GetString(parameters, "path");
        if (string.IsNullOrWhiteSpace(path))
        {
            _logger.LogWarning("runProcess action missing 'path' parameter.");
            return;
        }

        path = ResolvePath(module, path);
        var arguments = GetString(parameters, "arguments") ?? string.Empty;
        var workingDirectory = GetString(parameters, "workingDirectory");

        var startInfo = new ProcessStartInfo
        {
            FileName = path,
            Arguments = arguments,
            UseShellExecute = true
        };

        if (!string.IsNullOrWhiteSpace(workingDirectory))
        {
            startInfo.WorkingDirectory = ResolvePath(module, workingDirectory);
        }

        var process = StartProcess(startInfo);
        if (process is not null)
        {
            RegisterProcess(module, command, process);
        }
    }

    private void RunShellCommand(JsonObject parameters)
    {
        var command = GetString(parameters, "command");
        if (string.IsNullOrWhiteSpace(command))
        {
            _logger.LogWarning("shell action missing 'command' parameter.");
            return;
        }

        var shell = GetString(parameters, "shell") ?? "cmd.exe";
        var argsTemplate = GetString(parameters, "arguments") ?? "/C \"{0}\"";
        var arguments = string.Format(argsTemplate, command);

        var startInfo = new ProcessStartInfo
        {
            FileName = shell,
            Arguments = arguments,
            UseShellExecute = true
        };

        StartProcess(startInfo);
    }

    private void RunScript(CommandModule module, JsonObject parameters)
    {
        var scriptPath = GetString(parameters, "path");
        if (string.IsNullOrWhiteSpace(scriptPath))
        {
            _logger.LogWarning("script action missing 'path' parameter.");
            return;
        }

        var resolvedPath = ResolvePath(module, scriptPath);
        if (!File.Exists(resolvedPath))
        {
            _logger.LogWarning("Script file {Path} not found.", resolvedPath);
            return;
        }

        var extension = Path.GetExtension(resolvedPath).ToLowerInvariant();
        ProcessStartInfo startInfo;

        switch (extension)
        {
            case ".bat":
            case ".cmd":
                startInfo = new ProcessStartInfo("cmd.exe", $"/C \"{resolvedPath}\"")
                {
                    UseShellExecute = true
                };
                break;
            case ".ps1":
                startInfo = new ProcessStartInfo("powershell.exe", $"-ExecutionPolicy Bypass -File \"{resolvedPath}\"")
                {
                    UseShellExecute = true
                };
                break;
            default:
                startInfo = new ProcessStartInfo(resolvedPath)
                {
                    UseShellExecute = true
                };
                break;
        }

        StartProcess(startInfo);
    }

    private void SendKeys(JsonObject parameters)
    {
        var keysNode = parameters.TryGetPropertyValue("keys", out var node) ? node : null;
        if (keysNode is null)
        {
            _logger.LogWarning("sendKeys action missing 'keys' parameter.");
            return;
        }

        IEnumerable<string> tokens = keysNode switch
        {
            JsonArray array => array
                .Select(item => item?.GetValue<string>() ?? string.Empty)
                .Where(token => !string.IsNullOrWhiteSpace(token)),
            JsonValue value when value.TryGetValue<string>(out var token) =>
                token.Split('+', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries),
            _ => Array.Empty<string>()
        };

        _keyboardController.SendShortcut(tokens);
    }

    private void TypeText(JsonObject parameters)
    {
        var text = GetString(parameters, "text");
        if (string.IsNullOrEmpty(text))
        {
            _logger.LogWarning("typeText action missing 'text' parameter.");
            return;
        }

        _keyboardController.TypeText(text);
    }

    private void HandleMouseMove(JsonObject parameters, CancellationToken cancellationToken)
    {
        var dx = GetInt(parameters, "dx") ?? 0;
        var dy = GetInt(parameters, "dy") ?? 0;

        var continuous = GetBool(parameters, "continuous");
        if (continuous)
        {
            var interval = GetInt(parameters, "intervalMs") ?? 30;
            _mouseController.StartContinuousMove(dx, dy, interval, cancellationToken);
        }
        else
        {
            _mouseController.MoveBy(dx, dy);
        }
    }

    private void HandleMouseClick(JsonObject parameters)
    {
        var button = GetString(parameters, "button") ?? "left";
        var isDouble = GetBool(parameters, "double");
        _mouseController.Click(button, isDouble);
    }

    private void OpenUrl(JsonObject parameters)
    {
        var url = GetString(parameters, "url");
        if (string.IsNullOrWhiteSpace(url))
        {
            _logger.LogWarning("openUrl action missing 'url' parameter.");
            return;
        }

        var startInfo = new ProcessStartInfo
        {
            FileName = url,
            UseShellExecute = true
        };

        StartProcess(startInfo);
    }

    private void OpenWebView(CommandModule module, VoiceCommand command, JsonObject parameters)
    {
        _logger.LogInformation("OpenWebView called for module {ModuleId}, command {CommandId}", module.Id, command.Id);
        
        var url = GetString(parameters, "url");
        if (string.IsNullOrWhiteSpace(url))
        {
            _logger.LogWarning("openWebView action missing 'url' parameter.");
            return;
        }

        var topmost = GetBool(parameters, "top");
        var width = GetInt(parameters, "w") ?? 0;
        var height = GetInt(parameters, "h") ?? 0;
        var title = GetString(parameters, "title");
        var left = GetInt(parameters, "x") ?? GetInt(parameters, "left");
        var topPosition = GetInt(parameters, "y") ?? GetInt(parameters, "topPosition");

        _logger.LogInformation("Attempting to open WebView: URL={Url}, Topmost={Topmost}, Width={Width}, Height={Height}",
            url, topmost, width, height);

        try
        {
            if (System.Windows.Application.Current == null)
            {
                _logger.LogError("Application.Current is null - cannot create WebView window");
                return;
            }

            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                try
                {
                    _logger.LogInformation("Creating WebView window on UI thread...");
                    var windowLogger = _loggerFactory.CreateLogger<WebViewWindow>();
                    var webViewWindow = new WebViewWindow(
                        url,
                        title,
                        topmost,
                        width,
                        height,
                        left,
                        topPosition,
                        windowLogger);
                    
                    // Register window for tracking and cleanup
                    RegisterWebViewWindow(module, command, webViewWindow);
                    
                    // Handle window closed event
                    webViewWindow.WindowClosedEvent += (sender, e) =>
                    {
                        UnregisterWebViewWindow(module, command, webViewWindow);
                    };
                    
                    _logger.LogInformation("Showing WebView window...");
                    webViewWindow.Show();
                    
                    _logger.LogInformation("Opened WebView window for URL: {Url}, Topmost: {Topmost}, Size: {Width}x{Height}",
                        url, topmost, width, height);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Failed to create/show WebView window on UI thread for URL: {Url}", url);
                }
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to invoke Dispatcher for WebView window. URL: {Url}", url);
        }
    }

    private void OpenResource(CommandModule module, VoiceCommand command, JsonObject parameters)
    {
        var key = GetString(parameters, "key");
        if (string.IsNullOrWhiteSpace(key))
        {
            _logger.LogWarning("openResource action missing 'key' parameter.");
            return;
        }

        if (!_resourceRegistry.TryGet(key, out var resource))
        {
            _logger.LogWarning("Resource with key '{Key}' not found.", key);
            return;
        }

        var path = Environment.ExpandEnvironmentVariables(resource.Path);

        if (string.Equals(resource.Type, "Folder", StringComparison.OrdinalIgnoreCase))
        {
            var startInfo = new ProcessStartInfo("explorer.exe", $"\"{path}\"")
            {
                UseShellExecute = true
            };
            StartProcess(startInfo);
            
            // Track folder path for Explorer windows (can't track process directly)
            RegisterExplorerFolder(module, command, path);
            return;
        }

        var info = new ProcessStartInfo
        {
            FileName = path,
            UseShellExecute = true
        };
        var proc = StartProcess(info);
        if (proc is not null)
        {
            RegisterProcess(module, command, proc);
        }
    }

    private Task SpeakAsync(JsonObject parameters, CancellationToken cancellationToken)
    {
        var text = GetString(parameters, "text");
        if (string.IsNullOrWhiteSpace(text))
        {
            return Task.CompletedTask;
        }

        return _speechSynthesis.SpeakAsync(text, cancellationToken);
    }

    private async Task PauseAsync(JsonObject parameters, CancellationToken cancellationToken)
    {
        var milliseconds = GetInt(parameters, "milliseconds");
        if (milliseconds is null)
        {
            var seconds = GetDouble(parameters, "seconds");
            milliseconds = seconds is null ? 0 : (int)(seconds.Value * 1000);
        }

        if (milliseconds is null || milliseconds <= 0)
        {
            return;
        }

        await Task.Delay(milliseconds.Value, cancellationToken).ConfigureAwait(false);
    }

    private Process? StartProcess(ProcessStartInfo startInfo)
    {
        try
        {
            return Process.Start(startInfo);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to start process {FileName}.", startInfo.FileName);
            return null;
        }
    }

    private static string? GetString(JsonObject obj, string property)
    {
        if (obj.TryGetPropertyValue(property, out var node) &&
            node is JsonValue value &&
            value.TryGetValue<string>(out var result))
        {
            return result;
        }

        return null;
    }

    private static int? GetInt(JsonObject obj, string property)
    {
        if (obj.TryGetPropertyValue(property, out var node))
        {
            if (node is JsonValue value)
            {
                if (value.TryGetValue<int>(out var direct))
                {
                    return direct;
                }

                if (value.TryGetValue<string>(out var text) && int.TryParse(text, out var parsed))
                {
                    return parsed;
                }
            }
        }

        return null;
    }

    private static double? GetDouble(JsonObject obj, string property)
    {
        if (obj.TryGetPropertyValue(property, out var node))
        {
            if (node is JsonValue value)
            {
                if (value.TryGetValue<double>(out var direct))
                {
                    return direct;
                }

                if (value.TryGetValue<string>(out var text) && double.TryParse(text, out var parsed))
                {
                    return parsed;
                }
            }
        }

        return null;
    }

    private static bool GetBool(JsonObject obj, string property)
    {
        if (obj.TryGetPropertyValue(property, out var node))
        {
            if (node is JsonValue value)
            {
                if (value.TryGetValue<bool>(out var boolean))
                {
                    return boolean;
                }

                if (value.TryGetValue<string>(out var text) && bool.TryParse(text, out var parsed))
                {
                    return parsed;
                }
            }
        }

        return false;
    }

    private static string ResolvePath(CommandModule module, string path)
    {
        var expanded = Environment.ExpandEnvironmentVariables(path);

        if (Path.IsPathRooted(expanded))
        {
            return expanded;
        }

        var candidate = Path.GetFullPath(Path.Combine(module.BaseDirectory, expanded));
        if (File.Exists(candidate))
        {
            return candidate;
        }

        return expanded;
    }

    private void RegisterProcess(CommandModule module, VoiceCommand command, Process process)
    {
        var key = BuildCommandKey(module.Id, command.Id);

        process.EnableRaisingEvents = true;
        process.Exited += (_, _) => CleanupProcess(key, process);

        var list = _processes.GetOrAdd(key, _ => new List<Process>());
        lock (list)
        {
            list.Add(process);
        }

        _logger.LogInformation("Process {FileName} started for command {CommandName} ({CommandId}).",
            process.StartInfo.FileName, command.Name, command.Id);
    }

    private void StopProcess(CommandModule currentModule, VoiceCommand currentCommand, JsonObject parameters)
    {
        var commandId = GetString(parameters, "commandId");
        if (string.IsNullOrWhiteSpace(commandId))
        {
            commandId = currentCommand.Id;
        }

        var moduleId = GetString(parameters, "moduleId");
        if (string.IsNullOrWhiteSpace(moduleId))
        {
            moduleId = currentModule.Id;
        }

        var key = BuildCommandKey(moduleId, commandId);
        
        // Try to close a WebView window first
        if (TryCloseWebViewWindow(key, commandId, moduleId))
        {
            return;
        }
        
        // Try to close an Explorer folder
        if (TryCloseExplorerFolder(key, commandId, moduleId))
        {
            return;
        }
        
        // Otherwise, close a tracked process
        if (!_processes.TryGetValue(key, out var list))
        {
            _logger.LogWarning("No tracked processes for command {CommandId} ({ModuleId}).", commandId, moduleId);
            return;
        }

        Process? targetProcess = null;
        lock (list)
        {
            for (var i = list.Count - 1; i >= 0; i--)
            {
                var candidate = list[i];
                if (candidate.HasExited)
                {
                    list.RemoveAt(i);
                    continue;
                }

                targetProcess = candidate;
                list.RemoveAt(i);
                break;
            }
        }

        if (targetProcess is null)
        {            _logger.LogWarning("No active process instances for command {CommandId} ({ModuleId}).", commandId, moduleId);
            return;
        }

        try
        {
            _logger.LogInformation("Attempting to close process {FileName} (PID {Pid}) launched by command {CommandId}.",
                targetProcess.StartInfo.FileName, targetProcess.Id, commandId);

            var closedGracefully = false;
            if (targetProcess.MainWindowHandle != IntPtr.Zero)
            {
                closedGracefully = targetProcess.CloseMainWindow();
                if (closedGracefully)
                {
                    targetProcess.WaitForExit(2000);
                }
            }

            if (!targetProcess.HasExited)
            {
                if (!closedGracefully)
                {
                    _logger.LogInformation("Force terminating process {Pid}.", targetProcess.Id);
                }

                targetProcess.Kill(true);
                targetProcess.WaitForExit(2000);
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to stop process for command {CommandId} ({ModuleId}).", commandId, moduleId);
        }
    }

    private void StopAllProcesses(CommandModule currentModule, VoiceCommand currentCommand, JsonObject parameters)
    {
        var commandId = GetString(parameters, "commandId");
        var moduleId = GetString(parameters, "moduleId");

        // If both are specified, stop all processes for that specific command
        if (!string.IsNullOrWhiteSpace(commandId) && !string.IsNullOrWhiteSpace(moduleId))
        {
            var key = BuildCommandKey(moduleId, commandId);
            CloseAllWebViewWindows(key, commandId, moduleId);
            CloseAllExplorerFolders(key, commandId, moduleId);
            StopProcessesForKey(key, commandId, moduleId);
            return;
        }

        // Otherwise, stop ALL tracked processes, explorer folders, and webview windows
        _logger.LogInformation("Stopping all tracked processes, Explorer folders, and WebView windows...");
        
        var allKeys = _processes.Keys
            .Concat(_explorerFolders.Keys)
            .Concat(_webViewWindows.Keys)
            .Distinct()
            .ToList();
        var totalStopped = 0;

        foreach (var key in allKeys)
        {
            var parts = key.Split("::", StringSplitOptions.RemoveEmptyEntries);
            var moduleIdPart = parts.Length > 0 ? parts[0] : "unknown";
            var commandIdPart = parts.Length > 1 ? parts[1] : "unknown";
            
            var webViewsClosed = CloseAllWebViewWindows(key, commandIdPart, moduleIdPart);
            var explorersClosed = CloseAllExplorerFolders(key, commandIdPart, moduleIdPart);
            var processesStopped = StopProcessesForKey(key, commandIdPart, moduleIdPart);
            totalStopped += webViewsClosed + explorersClosed + processesStopped;
        }

        _logger.LogInformation("Stopped {Count} process(es) total.", totalStopped);
    }

    private int StopProcessesForKey(string key, string commandId, string moduleId)
    {
        if (!_processes.TryGetValue(key, out var list))
        {
            return 0;
        }

        List<Process> processesToStop;
        lock (list)
        {
            processesToStop = list.Where(p => !p.HasExited).ToList();
            list.Clear();
        }

        if (processesToStop.Count == 0)
        {
            return 0;
        }

        _logger.LogInformation("Stopping {Count} process(es) for command {CommandId} ({ModuleId})...",
            processesToStop.Count, commandId, moduleId);

        var stopped = 0;
        foreach (var process in processesToStop)
        {
            try
            {
                var closedGracefully = false;
                if (process.MainWindowHandle != IntPtr.Zero)
                {
                    closedGracefully = process.CloseMainWindow();
                    if (closedGracefully)
                    {
                        process.WaitForExit(2000);
                    }
                }

                if (!process.HasExited)
                {
                    process.Kill(true);
                    process.WaitForExit(2000);
                }

                stopped++;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to stop process {Pid}.", process.Id);
            }
        }

        return stopped;
    }

    private void CleanupProcess(string key, Process process)
    {
        if (!_processes.TryGetValue(key, out var list))
        {
            return;
        }

        lock (list)
        {
            list.Remove(process);
        }
    }

    private void RegisterExplorerFolder(CommandModule module, VoiceCommand command, string folderPath)
    {
        var key = BuildCommandKey(module.Id, command.Id);
        var normalizedPath = Path.GetFullPath(folderPath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

        var list = _explorerFolders.GetOrAdd(key, _ => new List<string>());
        lock (list)
        {
            list.Add(normalizedPath);
        }

        _logger.LogInformation("Explorer folder '{FolderPath}' registered for command {CommandName} ({CommandId}).",
            normalizedPath, command.Name, command.Id);
    }

    private bool TryCloseExplorerFolder(string key, string commandId, string moduleId)
    {
        if (!_explorerFolders.TryGetValue(key, out var folderList))
        {
            return false;
        }

        string? targetFolder = null;
        lock (folderList)
        {
            if (folderList.Count == 0)
            {
                return false;
            }

            targetFolder = folderList[folderList.Count - 1];
            folderList.RemoveAt(folderList.Count - 1);
        }

        if (targetFolder == null)
        {
            return false;
        }

        return CloseExplorerWindowByPath(targetFolder, commandId, moduleId);
    }

    private int CloseAllExplorerFolders(string key, string commandId, string moduleId)
    {
        if (!_explorerFolders.TryGetValue(key, out var folderList))
        {
            return 0;
        }

        List<string> foldersToClose;
        lock (folderList)
        {
            foldersToClose = new List<string>(folderList);
            folderList.Clear();
        }

        if (foldersToClose.Count == 0)
        {
            return 0;
        }

        _logger.LogInformation("Closing {Count} Explorer folder(s) for command {CommandId} ({ModuleId})...",
            foldersToClose.Count, commandId, moduleId);

        var closed = 0;
        foreach (var folder in foldersToClose)
        {
            if (CloseExplorerWindowByPath(folder, commandId, moduleId))
            {
                closed++;
            }
        }

        return closed;
    }

    private bool CloseExplorerWindowByPath(string folderPath, string commandId, string moduleId)
    {
        try
        {
            dynamic shellWindows = new ShellWindows();
            var normalizedPath = Path.GetFullPath(folderPath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

            _logger.LogInformation("Attempting to close Explorer window for folder '{FolderPath}' (command {CommandId})...",
                normalizedPath, commandId);

            int count = shellWindows.Count;
            for (int i = 0; i < count; i++)
            {
                dynamic window = null;
                try
                {
                    window = shellWindows.Item(i);
                    var locationUrl = window.LocationURL as string;
                    if (string.IsNullOrEmpty(locationUrl))
                    {
                        continue;
                    }

                    // Explorer windows have URLs like "file:///C:/Users/..."
                    if (!locationUrl.StartsWith("file:///", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    // Convert file:/// URL to path
                    var windowPath = new Uri(locationUrl).LocalPath;
                    var normalizedWindowPath = Path.GetFullPath(windowPath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

                    if (string.Equals(normalizedWindowPath, normalizedPath, StringComparison.OrdinalIgnoreCase))
                    {
                        _logger.LogInformation("Found matching Explorer window for '{FolderPath}'. Closing...", normalizedPath);
                        window.Quit();
                        if (window != null) Marshal.ReleaseComObject(window);
                        if (shellWindows != null) Marshal.ReleaseComObject(shellWindows);
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Error checking Explorer window.");
                }
                finally
                {
                    if (window != null)
                    {
                        try { Marshal.ReleaseComObject(window); } catch { /* ignore */ }
                    }
                }
            }

            if (shellWindows != null) Marshal.ReleaseComObject(shellWindows);
            _logger.LogWarning("No matching Explorer window found for folder '{FolderPath}'.", normalizedPath);
            return false;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to close Explorer window for folder '{FolderPath}'.", folderPath);
            return false;
        }
    }

    private static string BuildCommandKey(string moduleId, string commandId) =>
        $"{moduleId}::{commandId}";

    private void RegisterWebViewWindow(CommandModule module, VoiceCommand command, WebViewWindow window)
    {
        var key = BuildCommandKey(module.Id, command.Id);
        var list = _webViewWindows.GetOrAdd(key, _ => new List<WebViewWindow>());
        
        lock (list)
        {
            list.Add(window);
        }

        _logger.LogInformation("WebView window registered for command {CommandName} ({CommandId}).",
            command.Name, command.Id);
    }

    private void UnregisterWebViewWindow(CommandModule module, VoiceCommand command, WebViewWindow window)
    {
        var key = BuildCommandKey(module.Id, command.Id);
        
        if (!_webViewWindows.TryGetValue(key, out var list))
        {
            return;
        }

        lock (list)
        {
            list.Remove(window);
        }

        _logger.LogInformation("WebView window unregistered for command {CommandName} ({CommandId}).",
            command.Name, command.Id);
    }

    private bool TryCloseWebViewWindow(string key, string commandId, string moduleId)
    {
        if (!_webViewWindows.TryGetValue(key, out var windowList))
        {
            return false;
        }

        WebViewWindow? targetWindow = null;
        lock (windowList)
        {
            if (windowList.Count == 0)
            {
                return false;
            }

            targetWindow = windowList[windowList.Count - 1];
            windowList.RemoveAt(windowList.Count - 1);
        }

        if (targetWindow == null)
        {
            return false;
        }

        System.Windows.Application.Current.Dispatcher.Invoke(() =>
        {
            try
            {
                _logger.LogInformation("Closing WebView window for command {CommandId} ({ModuleId}).",
                    commandId, moduleId);
                targetWindow.Close();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to close WebView window for command {CommandId} ({ModuleId}).",
                    commandId, moduleId);
            }
        });

        return true;
    }

    private int CloseAllWebViewWindows(string key, string commandId, string moduleId)
    {
        if (!_webViewWindows.TryGetValue(key, out var windowList))
        {
            return 0;
        }

        List<WebViewWindow> windowsToClose;
        lock (windowList)
        {
            windowsToClose = new List<WebViewWindow>(windowList);
            windowList.Clear();
        }

        if (windowsToClose.Count == 0)
        {
            return 0;
        }

        _logger.LogInformation("Closing {Count} WebView window(s) for command {CommandId} ({ModuleId})...",
            windowsToClose.Count, commandId, moduleId);

        var closed = 0;
        foreach (var window in windowsToClose)
        {
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                try
                {
                    window.Close();
                    closed++;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Failed to close WebView window.");
                }
            });
        }

        return closed;
    }
}

// COM class for Shell Windows automation
[ComImport]
[Guid("9BA05972-F6A8-11CF-A442-00A0C90A8F39")]
[ClassInterface(ClassInterfaceType.None)]
internal class ShellWindows
{
}
