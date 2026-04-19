using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using TranslatorApp.Models;
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

    private void TaskGridRow_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is not DataGridRow row)
        {
            return;
        }

        row.IsSelected = true;
        row.Focus();

        var dataGrid = FindVisualParent<DataGrid>(row);
        if (dataGrid is not null)
        {
            dataGrid.SelectedItem = row.Item;
            dataGrid.CurrentItem = row.Item;
        }
    }

    private static T? FindVisualParent<T>(DependencyObject child) where T : DependencyObject
    {
        var current = VisualTreeHelper.GetParent(child);
        while (current is not null)
        {
            if (current is T target)
            {
                return target;
            }

            current = VisualTreeHelper.GetParent(current);
        }

        return null;
    }

    private async void ResumeTaskMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is MainViewModel viewModel && TaskGrid.SelectedItem is DocumentTranslationItem item)
        {
            await viewModel.ResumeTaskCommand.ExecuteAsync(item);
        }
    }

    private async void RemoveTaskMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is MainViewModel viewModel && TaskGrid.SelectedItem is DocumentTranslationItem item)
        {
            await viewModel.RemoveSelectedCommand.ExecuteAsync(item);
        }
    }

    private void OpenOutputDocumentMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is MainViewModel viewModel && TaskGrid.SelectedItem is DocumentTranslationItem item)
        {
            viewModel.OpenOutputDocumentCommand.Execute(item);
        }
    }

    private void OpenTranslationSettingsButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new TranslationSettingsWindow
        {
            Owner = this,
            DataContext = DataContext
        };

        dialog.ShowDialog();
    }

    private void LogTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            textBox.ScrollToEnd();
        }
    }
}
