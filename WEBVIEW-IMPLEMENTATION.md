# WebView Implementation Guide

## Overview
This document describes the implementation of the new WebView functionality in VoicePilot, which allows opening URLs in a custom application window using WebView2 instead of the default browser.

## Changes Made

### 1. Added WebView2 NuGet Package
- **File**: `VoicePilot.App.csproj`
- **Change**: Added `Microsoft.Web.WebView2` version 1.0.2792.45 package reference

### 2. Created WebViewWindow Component
Created two new files for the WebView window:

#### WebViewWindow.xaml
- Custom WPF window with dark theme
- Title bar with navigation controls (Back, Forward, Refresh, Close)
- WebView2 control for displaying web content
- Supports both windowed and fullscreen modes
- Topmost window option

#### WebViewWindow.xaml.cs
- Window logic and event handlers
- Parameters:
  - `url`: The URL to open
  - `title`: Optional custom caption for the window
  - `topmost`: Whether window stays on top of all others
  - `width`: Window width (0 for fullscreen)
  - `height`: Window height (0 for fullscreen)
  - `x` / `left`: Manual X coordinate (optional)
  - `y` / `topPosition`: Manual Y coordinate (optional)
- Features:
  - Async WebView2 initialization
  - Navigation controls (back, forward, refresh)
  - Draggable title bar
  - Double-click title bar to maximize/restore
  - WindowClosedEvent for tracking
  - Optional manual placement when coordinates are provided

### 3. Updated ActionExecutor Service
- **File**: `ActionExecutor.cs`
- **Changes**:
  - Added `_webViewWindows` dictionary to track open WebView windows
  - Implemented `OpenWebView()` method to handle "openWebView" action type
  - Added WebView window registration/unregistration methods
  - Integrated WebView window closing into `stopProcess` and `stopAllProcesses` commands
  - WebView windows are now tracked like other processes and can be closed using voice commands

### 4. Created YouTube WebView Module
- **File**: `Modules/youtube.webview-1.0.0/manifest.json`
- **Commands**:
  1. **Open YouTube WebView**:
     - Phrases: "открой ютуб в окне", "включи ютуб в окне"
     - Opens YouTube in 840x480 window, always on top
  2. **Close YouTube WebView**:
     - Phrases: "закрой ютуб в окне", "выключи ютуб в окне"
     - Closes the last opened YouTube WebView window

## How to Use

### Creating a WebView Command

Create a module manifest with the following structure:

```json
{
  "id": "your.module.id",
  "name": "Module Name",
  "version": "1.0.0",
  "commands": [
    {
      "id": "command_id",
      "name": "Command Name",
      "phrases": ["voice phrase 1", "voice phrase 2"],
      "actions": [
        {
          "type": "openWebView",
          "parameters": {
            "url": "https://example.com",
            "title": "Custom Tool",
            "top": true,
            "w": 840,
            "h": 480,
            "x": 120,
            "y": 80
          }
        }
      ]
    }
  ]
}
```

### Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `url` | string | Yes | The URL to open in WebView |
| `top` | boolean | No | If true, window stays on top of all others |
| `title` | string | No | Custom caption for the window header |
| `w` | integer | No | Window width (0 or omitted = fullscreen) |
| `h` | integer | No | Window height (0 or omitted = fullscreen) |
| `x` / `left` | integer | No | Manual X coordinate (forces windowed mode) |
| `y` / `topPosition` | integer | No | Manual Y coordinate (forces windowed mode) |

### Fullscreen Mode
- Set `w` and `h` to 0 or omit them for fullscreen mode
- Example: `"w": 0, "h": 0` or just omit both parameters

### Closing WebView Windows

#### Close Last Opened Window
```json
{
  "type": "stopProcess",
  "parameters": {
    "commandId": "your_command_id"
  }
}
```

#### Close All Windows for a Command
```json
{
  "type": "stopAllProcesses",
  "parameters": {
    "commandId": "your_command_id",
    "moduleId": "your.module.id"
  }
}
```

#### Close All Windows System-Wide
```json
{
  "type": "stopAllProcesses",
  "parameters": {}
}
```

## Technical Details

### Process Tracking
WebView windows are tracked similar to other processes:
- Registered when opened with `RegisterWebViewWindow()`
- Tracked by module ID and command ID
- Unregistered when closed via `WindowClosedEvent`
- Can be closed programmatically via `stopProcess` and `stopAllProcesses`

### Window Management
- Windows are created on the WPF UI thread using `Dispatcher.Invoke()`
- Each window instance is tracked in `_webViewWindows` dictionary
- Windows are closed gracefully with proper cleanup
- COM threading model is handled correctly

### Integration with Existing System
The WebView functionality integrates seamlessly with the existing process management:
1. `stopProcess` tries to close WebView windows first, then Explorer folders, then processes
2. `stopAllProcesses` closes all tracked resources including WebView windows
3. WebView windows work like `openResource` - they can be closed using the same commands

## Requirements

### WebView2 Runtime
Users must have Microsoft Edge WebView2 Runtime installed:
- Typically pre-installed on Windows 10/11
- Download from: https://developer.microsoft.com/microsoft-edge/webview2/
- If missing, the app will show an error message with instructions

## Example Use Cases

1. **YouTube in Small Window**: 840x480 window, always on top
2. **Documentation Browser**: Fullscreen WebView for reading docs
3. **Web Tools**: Open web-based tools in dedicated windows
4. **Media Players**: Web-based media players in custom-sized windows
5. **Dashboard Monitors**: Always-on-top dashboards

## Future Enhancements

Possible improvements:
- Add zoom controls
- Cookie/session persistence
- Custom user agent
- JavaScript injection support
- Download handling
- Custom context menus
- Multiple monitor support

## Comparison: openUrl vs openWebView

| Feature | openUrl | openWebView |
|---------|---------|-------------|
| Opens in | Default browser | App window |
| Window control | No | Yes (size, topmost) |
| Process tracking | No | Yes |
| Voice close | No | Yes |
| Navigation controls | Browser's | Custom title bar |
| Fullscreen support | Browser dependent | Built-in |
| Multiple instances | Browser tabs | Separate windows |

## Testing

To test the implementation:

1. Build the project: `dotnet build`
2. Run the application
3. Say: "открой ютуб в окне" (or equivalent phrase)
4. Verify window opens at 840x480, stays on top
5. Test navigation controls (back, forward, refresh)
6. Say: "закрой ютуб в окне" to close
7. Verify window closes properly

## Troubleshooting

### WebView2 Not Loading
- Ensure WebView2 Runtime is installed
- Check for firewall/antivirus blocking
- Verify internet connection for online content

### Window Not Staying on Top
- Check `top` parameter is set to `true`
- Verify no other always-on-top windows conflict

### Build Errors
- Ensure `Microsoft.Web.WebView2` NuGet package is restored
- Clean and rebuild: `dotnet clean && dotnet build`

---

**Implementation Date**: November 9, 2025
**Version**: 1.0.0
