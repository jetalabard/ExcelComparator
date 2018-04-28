using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Windows;
using System.Windows.Media.Imaging;
using Tools;

namespace ExcelComparator
{
    #pragma warning disable S110 // Inheritance tree of classes should not be too deep
    /// <summary>
    /// Logique d'interaction pour MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    #pragma warning restore S110 // Inheritance tree of classes should not be too deep
    {

        private ApplicationController controler;

        public MainWindow()
        {
            InitializeComponent();
            Uri iconUri = new Uri(GetImageUrl.Icon(), UriKind.RelativeOrAbsolute);
            Icon = BitmapFrame.Create(iconUri);
            controler = new ApplicationController();

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            file1.Text = controler.OpenDialogAndGetFileName();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            file2.Text = controler.OpenDialogAndGetFileName();
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            if (controler.VerifyFieldFile(file1.Text, file2.Text))
            {
                string sheetField = string.Empty;
                string columsField = string.Empty;
                string sheet2Field = string.Empty;
                if (!string.IsNullOrEmpty(columns.Text))
                {
                    columsField = columns.Text;
                }
                if (!string.IsNullOrEmpty(sheet.Text))
                {
                    sheetField = sheet.Text;
                }
                if (!string.IsNullOrEmpty(sheet2.Text))
                {
                    sheet2Field = sheet2.Text;
                }
                controler.Compare(file1.Text, sheetField, file2.Text, sheet2Field, columsField);
            }
        }


    }
}
