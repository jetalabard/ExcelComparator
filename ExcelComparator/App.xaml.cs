using System.Threading.Tasks;
using System.Windows;

namespace ExcelComparator
{
    /// <summary>
    /// Logique d'interaction pour App.xaml
    /// </summary>
    public partial class App : Application
    {
        /// <summary>
        /// écran de chargement
        /// </summary>
        private Tools.SplashScreen splashScreen;

        /// <summary>
        /// Démarre l'application en gérant le lancement Console ou Interface
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void App_Startup(object sender, StartupEventArgs e)
        {

            MainWindow = new MainWindow();
            splashScreen = new Tools.SplashScreen(@"Image\Splash\Splash.png");
            splashScreen.Show();
        }

        /// <summary>
        /// démarre l'application interface après avoir affiché le dialogue de chargement
        /// </summary>
        /// <param name="e"></param>
        protected override async void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            await Task.Delay(500);
            splashScreen.SetProgress("", 3, 4);
            MainWindow.Show();
            splashScreen.Close();
            splashScreen = null;
        }

    }
}
