using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using Application = System.Windows.Application;
using CommunityToolkit.Mvvm.Messaging;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using VoicePilot.App.Configuration;
using VoicePilot.App.Input;
using VoicePilot.App.Messaging;
using VoicePilot.App.Modules;
using VoicePilot.App.Recognition;
using VoicePilot.App.Resources;
using VoicePilot.App.Services;
using VoicePilot.App.ViewModels;

namespace VoicePilot.App;

public partial class App : Application
{
    private IHost? _host;

    protected override async void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        _host = Host.CreateDefaultBuilder()
            .UseContentRoot(AppContext.BaseDirectory)
            .ConfigureAppConfiguration((context, builder) =>
            {
                builder.Sources.Clear();

                builder.AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
                builder.AddJsonFile(Path.Combine("config", "appsettings.json"), optional: true, reloadOnChange: true);
                builder.AddEnvironmentVariables(prefix: "VOICEPILOT_");
            })
            .ConfigureLogging(logging =>
            {
                logging.ClearProviders();
                logging.AddConsole();
            })
            .ConfigureServices(ConfigureServices)
            .Build();

        await _host.StartAsync().ConfigureAwait(false);

        var mainWindow = _host.Services.GetRequiredService<MainWindow>();
        MainWindow = mainWindow;
        mainWindow.Show();
    }

    private static void ConfigureServices(HostBuilderContext context, IServiceCollection services)
    {
        services.AddSingleton<IMessenger>(WeakReferenceMessenger.Default);

        services.AddOptions<SpeechOptions>()
            .Bind(context.Configuration.GetSection("speech"))
            .PostConfigure(options =>
            {
                if (string.IsNullOrWhiteSpace(options.ModelPath))
                {
                    options.ModelPath = Path.Combine("models", "vosk-model-small-ru-0.22");
                }
            });

        services.AddOptions<PathOptions>()
            .Bind(context.Configuration.GetSection("paths"))
            .PostConfigure(options =>
            {
                options.Resolve(AppContext.BaseDirectory);
                Directory.CreateDirectory(options.ModulesDirectory);
                Directory.CreateDirectory(options.ModelsDirectory);

                var systemCommandsDir = Path.GetDirectoryName(options.SystemCommandsFile);
                if (!string.IsNullOrWhiteSpace(systemCommandsDir))
                {
                    Directory.CreateDirectory(systemCommandsDir);
                }

                var resourcesDir = Path.GetDirectoryName(options.ResourcesFile);
                if (!string.IsNullOrWhiteSpace(resourcesDir))
                {
                    Directory.CreateDirectory(resourcesDir);
                }
            });

        services.AddSingleton<MainViewModel>();
        services.AddSingleton<MainWindow>();

        services.AddSingleton<ModuleManifestParser>();
        services.AddSingleton<ModuleManager>();
        services.AddSingleton<CommandRouter>();
        services.AddSingleton<ActionExecutor>();
        services.AddSingleton<SystemCommandService>();
        services.AddSingleton<SpeechSynthesisService>();
        services.AddSingleton<ResourceRegistry>();
        services.AddSingleton<MouseController>();
        services.AddSingleton<KeyboardController>();

        services.AddHostedService<SpeechRecognitionHostedService>();
    }

    protected override async void OnExit(ExitEventArgs e)
    {
        if (_host is not null)
        {
            await _host.StopAsync(TimeSpan.FromSeconds(5)).ConfigureAwait(false);
            _host.Dispose();
        }

        base.OnExit(e);
    }
}
