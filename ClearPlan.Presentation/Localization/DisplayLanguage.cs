using System;
using System.ComponentModel;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Markup;
using ClearPlan.Core.Localization;

namespace ClearPlan.Presentation.Localization
{
    public sealed class DisplayLanguage : INotifyPropertyChanged
    {
        public static DisplayLanguage Instance { get; } = new DisplayLanguage();
        private DisplayLanguage()
        {
            ReviewLanguage.LoadPreference();
            ReviewLanguage.Changed += (s,e) => { var h=PropertyChanged; if(h!=null) h(this,new PropertyChangedEventArgs("Code")); };
        }
        public event PropertyChangedEventHandler PropertyChanged;
        public string[] Codes { get { return new[] {"de","en"}; } }
        public string Code { get { return ReviewLanguage.Code; } set { ReviewLanguage.Code=value; ReviewLanguage.SavePreference(); } }
    }
    /// <summary>Explicit opt-in for display strings; identity bindings are not localized.</summary>
    public sealed class TranslateExtension : MarkupExtension
    {
        public object Value { get; set; }
        public string Path { get; set; }
        public override object ProvideValue(IServiceProvider provider)
        {
            var binding=new MultiBinding { Mode=BindingMode.OneWay, Converter=TranslationConverter.Instance };
            binding.Bindings.Add(Path == null ? new Binding { Source=Value ?? string.Empty } : new Binding(Path) { Mode=BindingMode.OneWay });
            binding.Bindings.Add(new Binding("Code") { Source=DisplayLanguage.Instance });
            return binding.ProvideValue(provider);
        }
        private sealed class TranslationConverter : IMultiValueConverter
        {
            public static readonly TranslationConverter Instance=new TranslationConverter();
            public object Convert(object[] values,Type target,object parameter,CultureInfo culture)
            {
                if(values.Length==0 || values[0]==DependencyProperty.UnsetValue) return string.Empty;
                return ReviewLanguage.Display(values[0] as string ?? System.Convert.ToString(values[0],culture));
            }
            public object[] ConvertBack(object value,Type[] targets,object parameter,CultureInfo culture) { throw new NotSupportedException(); }
        }
    }
}
