using System.ComponentModel;
using System.Diagnostics;
using System.Text;

namespace ExcelPassword.Models
{
    class TextBoxTraceListener : TraceListener, INotifyPropertyChanged
    {
        public readonly StringBuilder builder;

        public TextBoxTraceListener()
        {
            this.builder = new StringBuilder();
        }

        public string Trace
        {
            get { return this.builder.ToString(); }
        }

        public override void Write(string message)
        {
            this.builder.Append(message);
            OnPropertyChanged("logging");
        }

        public override void WriteLine(string message)
        {
            this.builder.AppendLine(message);
            OnPropertyChanged("logging");
        }
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(name));
            }
        }

    }
}
