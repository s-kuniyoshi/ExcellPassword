using ExcelPassword.Models;
using System;
using System.Windows.Data;

namespace ExcelPassword.ViewModels
{
    class LoggingConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            string target = ((TextBoxTraceListener)value).Trace;

            return target;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
