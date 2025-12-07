using System;
using System.IO;
using System.Windows;
using System.Windows.Input;
using Microsoft.Extensions.Logging;
using Microsoft.Web.WebView2.Core;

namespace VoicePilot.App.Windows;

public partial class WebViewWindow : Window
{
    private readonly ILogger<WebViewWindow>? _logger;
    private readonly string _url;
    private readonly string? _title;
    private readonly bool _topmost;
    private readonly int _width;
    private readonly int _height;
    private readonly int? _left;
    private readonly int? _top;
    private bool _webViewDisposed;

    public event EventHandler? WindowClosedEvent;

    public WebViewWindow(
        string url,
        string? title = null,
        bool topmost = false,
        int width = 0,
        int height = 0,
        int? left = null,
        int? top = null,
        ILogger<WebViewWindow>? logger = null)
    {
        _url = url;
        _title = title;
        _topmost = topmost;
        _width = width;
        _height = height;
        _left = left;
        _top = top;
        _logger = logger;

        InitializeComponent();
        
        ConfigureWindow();
    }

    private void ConfigureWindow()
    {
        Topmost = _topmost;

        var hasExplicitSize = _width > 0 && _height > 0;
        if (hasExplicitSize)
        {
            Width = _width;
            Height = _height;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            WindowState = WindowState.Normal;
        }
        else
        {
            WindowState = WindowState.Maximized;
        }

        if (_left.HasValue || _top.HasValue)
        {
            WindowStartupLocation = WindowStartupLocation.Manual;
            WindowState = WindowState.Normal;

            if (_left.HasValue)
            {
                Left = _left.Value;
            }

            if (_top.HasValue)
            {
                Top = _top.Value;
            }
        }

        Title = string.IsNullOrWhiteSpace(_title) ? $"Web Viewer - {_url}" : _title;
        UrlTextBlock.Text = _url;
    }

    private async void Window_Loaded(object sender, RoutedEventArgs e)
    {
        try
        {
            // Create a custom user data folder for WebView2 to avoid permission issues
            var userDataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "VoicePilot",
                "WebView2",
                Guid.NewGuid().ToString()
            );
            
            // Ensure the directory exists
            Directory.CreateDirectory(userDataFolder);
            
            // Create WebView2 environment with custom user data folder
            var environment = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
            await WebView.EnsureCoreWebView2Async(environment);
            WebView.CoreWebView2.Navigate(_url);

            WebView.CoreWebView2.SourceChanged += (s, args) =>
            {
                Dispatcher.Invoke(() =>
                {
                    UrlTextBlock.Text = WebView.CoreWebView2.Source;
                });
            };

            _logger?.LogInformation("WebView loaded successfully for URL: {Url} with user data folder: {UserDataFolder}", _url, userDataFolder);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to initialize WebView2 for URL: {Url}", _url);
            System.Windows.MessageBox.Show(
                $"Failed to load WebView2: {ex.Message}\n\nPlease ensure WebView2 Runtime is installed.",
                "WebView Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            Close();
        }
    }

    private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ClickCount == 2)
        {
            WindowState = WindowState == WindowState.Maximized 
                ? WindowState.Normal 
                : WindowState.Maximized;
        }
        else
        {
            DragMove();
        }
    }

    private void BackButton_Click(object sender, RoutedEventArgs e)
    {
        if (WebView?.CoreWebView2?.CanGoBack == true)
        {
            WebView.CoreWebView2.GoBack();
        }
    }

    private void ForwardButton_Click(object sender, RoutedEventArgs e)
    {
        if (WebView?.CoreWebView2?.CanGoForward == true)
        {
            WebView.CoreWebView2.GoForward();
        }
    }

    private void RefreshButton_Click(object sender, RoutedEventArgs e)
    {
        WebView?.CoreWebView2?.Reload();
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
    {
        _logger?.LogInformation("WebView window closing for URL: {Url}", _url);
        CleanupWebView();
        WindowClosedEvent?.Invoke(this, EventArgs.Empty);
    }

    private void CleanupWebView()
    {
        if (_webViewDisposed)
        {
            return;
        }

        _webViewDisposed = true;

        CoreWebView2? core = null;
        try
        {
            core = WebView?.CoreWebView2;
        }
        catch (Exception ex)
        {
            _logger?.LogDebug(ex, "Failed to get CoreWebView2 instance for cleanup (URL: {Url}).", _url);
        }

        if (core is not null)
        {
            try
            {
                core.Stop();
            }
            catch (Exception ex)
            {
                _logger?.LogDebug(ex, "CoreWebView2.Stop() threw during cleanup (URL: {Url}).", _url);
            }

            try
            {
                core.Navigate("about:blank");
            }
            catch (Exception ex)
            {
                _logger?.LogDebug(ex, "CoreWebView2.Navigate(\"about:blank\") threw during cleanup (URL: {Url}).", _url);
            }
        }

        try
        {
            if (WebView is not null)
            {
                WebView.Source = null;
                WebView.Dispose();
            }
        }
        catch (Exception ex)
        {
            _logger?.LogWarning(ex, "Failed to dispose WebView control for URL: {Url}", _url);
        }
    }
}
