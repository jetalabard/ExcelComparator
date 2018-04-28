using System;
using System.Windows;
using System.Windows.Media.Imaging;

namespace Tools
{
    /// <summary>
    /// Logique d'interaction pour LoadingWindow.xaml
    /// </summary>
    #pragma warning disable S110 // Inheritance tree of classes should not be too deep
    public partial class LoadingWindow : Window
    #pragma warning restore S110 // Inheritance tree of classes should not be too deep
    {
       
        /// <summary>
        /// est-ce que le dialogue peut être fermé ?
        /// </summary>
        public bool Closable { get; private set; }

        /// <summary>
        /// constructeur
        /// </summary>
        /// <param name="closable"></param>
        public LoadingWindow(bool closable)
        {
            InitializeComponent();
            Closable = closable;
            Uri iconUri = new Uri(GetImageUrl.Icon(), UriKind.RelativeOrAbsolute);
            Icon = BitmapFrame.Create(iconUri);
        }

        public new void Close()
        {
            if (Closable)
            {
                base.Close();
            }
        }
        /// <summary>
        /// force à fermer le dialogue
        /// </summary>
        public void ForceClose()
        {
            base.Close();
        }
    }
}
