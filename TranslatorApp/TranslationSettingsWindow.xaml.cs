using System.Windows;

namespace TranslatorApp;

public partial class TranslationSettingsWindow : Window
{
    public TranslationSettingsWindow()
    {
        InitializeComponent();
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
    }
}
