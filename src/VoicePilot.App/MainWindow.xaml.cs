using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using VoicePilot.App.ViewModels;

namespace VoicePilot.App;

public partial class MainWindow : Window
{
    private bool _isMinimized;
    private Window? _miniWindow;

    public MainWindow(MainViewModel viewModel)
    {
        InitializeComponent();
        DataContext = viewModel;
        Loaded += MainWindow_OnLoaded;
    }

    private async void MainWindow_OnLoaded(object sender, RoutedEventArgs e)
    {
        Loaded -= MainWindow_OnLoaded;

        if (DataContext is MainViewModel vm)
        {
            await vm.EnsureModulesLoadedAsync();
        }
    }

    private void ModuleDataGrid_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (DataContext is not MainViewModel vm)
        {
            return;
        }

        if (vm.SelectedCommandItem is null)
        {
            return;
        }

        vm.OpenCommandEditor(vm.SelectedCommandItem);
    }

    private void StatusImage_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (!_isMinimized)
        {
            MinimizeToImage();
        }
        else
        {
            RestoreFromImage();
        }
    }

    private void MinimizeToImage()
    {
        try
        {
            if (DataContext is not MainViewModel vm)
            {
                return;
            }

            // Calculate position so mini window's top-right corner matches main window's top-right corner
            var miniWindowWidth = 120.0;
            var miniWindowHeight = 120.0;
            var mainWindowRight = Left + Width;
            var mainWindowTop = Top;
            
            var miniWindowLeft = mainWindowRight - miniWindowWidth;
            var miniWindowTop = mainWindowTop;

            // Hide main window
            Hide();

            // Create mini window
            _miniWindow = new Window
            {
                Width = miniWindowWidth,
                Height = miniWindowHeight,
                WindowStyle = WindowStyle.None,
                ResizeMode = ResizeMode.NoResize,
                AllowsTransparency = true,
                Background = System.Windows.Media.Brushes.Transparent,
                Topmost = true,
                ShowInTaskbar = false,
                Left = miniWindowLeft,
                Top = miniWindowTop
            };

            // Create image for mini window
            var image = new System.Windows.Controls.Image
            {
                Width = 96,
                Height = 96,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
                VerticalAlignment = System.Windows.VerticalAlignment.Center,
                Cursor = System.Windows.Input.Cursors.Hand
            };

            // Bind to the same image source
            var binding = new System.Windows.Data.Binding("ListeningIndicatorImage")
            {
                Source = vm
            };
            image.SetBinding(System.Windows.Controls.Image.SourceProperty, binding);

            // Add click handler to restore
            image.MouseLeftButtonDown += (s, e) => RestoreFromImage();

            // Add border
            var border = new Border
            {
                Width = 110,
                Height = 110,
                CornerRadius = new CornerRadius(12),
                BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xE0, 0xE5, 0xF2)),
                BorderThickness = new Thickness(2),
                Background = System.Windows.Media.Brushes.White,
                Child = image,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
                VerticalAlignment = System.Windows.VerticalAlignment.Center
            };

            _miniWindow.Content = border;
            _miniWindow.Show();

            _isMinimized = true;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error minimizing window: {ex.Message}");
            System.Windows.MessageBox.Show($"Ошибка при сворачивании окна: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void RestoreFromImage()
    {
        try
        {
            // Close mini window
            if (_miniWindow != null)
            {
                _miniWindow.Close();
                _miniWindow = null;
            }

            // Show main window
            Show();
            Activate();

            _isMinimized = false;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error restoring window: {ex.Message}");
            System.Windows.MessageBox.Show($"Ошибка при восстановлении окна: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}
