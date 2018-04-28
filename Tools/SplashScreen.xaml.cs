using System;
using System.Windows;
using System.Windows.Media.Imaging;

namespace Tools
{
    #pragma warning disable S110 // Inheritance tree of classes should not be too deep
    /// <summary>
    /// Logique d'interaction pour SplashScreenWindow.xaml
    /// </summary>
    public partial class SplashScreen : Window
    #pragma warning restore S110 // Inheritance tree of classes should not be too deep
    {
        /// <summary>
        /// constructeur
        /// </summary>
        /// <param name="imagePath"></param>
        public SplashScreen(string imagePath)
        {
            imagePath = System.IO.Path.GetFullPath(imagePath);

            InitializeComponent();

            Application.Current.Activated += OnAppActivated;
            Application.Current.Deactivated += OnAppDeactivated;
            Application.Current.MainWindow.Activated += OnMainWindowActivated;

            BitmapImage image;
            back.Source = image = new BitmapImage(new Uri(imagePath));

            Width = image.PixelWidth;
            Height = image.PixelHeight;
        }

        #region Activate
        private void OnAppActivated(object sender, EventArgs e)
        {
            Topmost = true;
        }

        private void OnMainWindowActivated(object sender, EventArgs e)
        {
            Topmost = true;
        }

        protected override void OnActivated(EventArgs e)
        {
            base.OnActivated(e);
            Topmost = true;
        }
        #endregion

        #region Desactivate
        private void OnAppDeactivated(object sender, EventArgs e)
        {
            Topmost = false;
        }

        protected override void OnDeactivated(EventArgs e)
        {
            base.OnDeactivated(e);
            Topmost = false;
        }
        #endregion

        protected override void OnClosed(EventArgs e)
        {
            Application.Current.Activated -= OnAppActivated;
            Application.Current.Deactivated -= OnAppDeactivated;
            Application.Current.MainWindow.Activated -= OnMainWindowActivated;

            base.OnClosed(e);
        }

        #region progress
        public void SetProgress(double value)
        {
            progress.IsIndeterminate = false;
            progress.Value = value;
        }

        public void SetProgress(double value, double maximum)
        {
            progress.IsIndeterminate = false;
            progress.Value = value;
            progress.Maximum = maximum;
        }
        public void SetProgress(string message, double value)
        {
            label.Text = message;
            progress.IsIndeterminate = false;
            progress.Value = value;
        }

        public void SetProgress(string message, double value, double maximum)
        {
            label.Text = message;
            progress.IsIndeterminate = false;
            progress.Value = value;
            progress.Maximum = maximum;
        }
        #endregion
    }
}
