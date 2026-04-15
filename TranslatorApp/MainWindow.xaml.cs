using System.Windows;
using TranslatorApp.ViewModels;

namespace TranslatorApp;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }

    private async void Window_Loaded(object sender, RoutedEventArgs e)
    {
        if (DataContext is MainViewModel viewModel)
        {
            await viewModel.LoadedCommand.ExecuteAsync(null);
        }
    }
}
