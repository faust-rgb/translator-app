using System.Globalization;
using System.Windows.Data;
using TranslatorApp.Models;

namespace TranslatorApp.Converters;

public sealed class DocumentStatusToTextConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture) =>
        value is DocumentStatus status
            ? status switch
            {
                DocumentStatus.Pending => "等待中",
                DocumentStatus.Running => "运行中",
                DocumentStatus.Paused => "已暂停",
                DocumentStatus.Completed => "已完成",
                DocumentStatus.Failed => "失败",
                DocumentStatus.Stopped => "已停止",
                _ => "未知"
            }
            : "未知";

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) =>
        Binding.DoNothing;
}
