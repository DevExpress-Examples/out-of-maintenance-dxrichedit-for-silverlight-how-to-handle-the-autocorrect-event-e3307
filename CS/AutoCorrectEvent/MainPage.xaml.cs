using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using System.IO;
using System.Reflection;
using DevExpress.XtraRichEdit;
using System.Globalization;
using System.Windows.Media.Imaging;
using DevExpress.XtraRichEdit.Utils;
using DevExpress.Office.Utils;
using DevExpress.Office.NumberConverters;

namespace AutoCorrectEvent
{
    public partial class MainPage : UserControl
    {
        public MainPage()
        {
            InitializeComponent();
        }
        private void richEditControl1_Loaded(object sender, RoutedEventArgs e)
        {
            richEditControl1.ApplyTemplate();
            richEditControl1.CreateNewDocument();
            richEditControl1.AutoCorrect += new AutoCorrectEventHandler(richEditControl1_AutoCorrect);
        }

        private OfficeImage CreateImageFromResx(string name)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            Stream stream = assembly.GetManifestResourceStream("AutoCorrectEvent.Images." + name);
            OfficeImage im = OfficeImage.CreateImage(stream);          
            return im;
        }
        #region #autocorrect
        private void richEditControl1_AutoCorrect(object sender, DevExpress.XtraRichEdit.AutoCorrectEventArgs e)
        {
            AutoCorrectInfo info = e.AutoCorrectInfo;
            e.AutoCorrectInfo = null;

            if (info.Text.Length <= 0)
                return;
            for (; ; ) {
                if (!info.DecrementStartPosition())
                    return;

                if (IsSeparator(info.Text[0]))
                    return;

                if (info.Text[0] == '$') {
                    info.ReplaceWith = CreateImageFromResx("dollar_pic.png");
                    e.AutoCorrectInfo = info;
                    return;
                }

                if (info.Text[0] == '%') {
                    string replaceString = CalculateFunction(info.Text);
                    if (!String.IsNullOrEmpty(replaceString)) {
                        info.ReplaceWith = replaceString;
                        e.AutoCorrectInfo = info;
                    }
                    return;
                }
            }
        }
        #endregion #autocorrect
        string CalculateFunction(string name)
        {
            name = name.ToLower();

            if (name.Length > 2 && name[0] == '%' && name.EndsWith("%")) {
                int value;
                if (Int32.TryParse(name.Substring(1, name.Length - 2), out value)) {
                    OrdinalBasedNumberConverter converter = OrdinalBasedNumberConverter.CreateConverter(NumberingFormat.CardinalText, LanguageId.English);
                    return converter.ConvertNumber(value);
                }
            }

            switch (name) {
                case "%date%":
                    return DateTime.Now.ToString(CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern);
                case "%time%":
                    return DateTime.Now.ToString(CultureInfo.CurrentCulture.DateTimeFormat.ShortTimePattern);
                default:
                    return String.Empty;
            }
        }
        bool IsSeparator(char ch)
        {
            return ch != '%' && (ch == '\r' || ch == '\n' || Char.IsPunctuation(ch) || Char.IsSeparator(ch) || Char.IsWhiteSpace(ch));
        }

    }
}
