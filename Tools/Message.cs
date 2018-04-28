using BespokeFusion;
using System.Windows;
using System.Windows.Media;

namespace Tools
{
    /// <summary>
    /// centralise la gestion des messagebox
    /// </summary>
    public static class Message
    {
       
        /// <summary>
        /// Affiche une boîte de message avec le texte
        /// </summary>
        /// <param name="text"></param>
        public static MessageBoxResult Show(string text)
        {
            var message = NormalMessageBox(text, "Erreur", "Ok");
            message.Show();
            return message.Result;
        }

        /// <summary>
        /// Affiche une boîte de message cancel avec gestion du résultat
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public static MessageBoxResult ShowWithCancel(string text,string textButtonOk)
        {
            var message = CancelMessageBox(text, "Erreur", textButtonOk);
            message.Show();
            return message.Result;
        }


        /// <summary>
        /// retourne la popup avec un message, le bouton Ok et le bouton annuler
        /// </summary>
        /// <param name="text"></param>
        /// <param name="title"></param>
        /// <param name="textButtonOk"></param>
        /// <returns></returns>
        private static CustomMaterialMessageBox CancelMessageBox(string text, string title, string textButtonOk)
        {
            CustomMaterialMessageBox messageBox = MessageBox(text,title,textButtonOk);
            messageBox.BtnCancel.Visibility = Visibility.Visible;
            return messageBox;
        }

        /// <summary>
        /// retourne la popup normale, avec un message et le bouton Ok
        /// </summary>
        /// <param name="text"></param>
        /// <param name="title"></param>
        /// <param name="textButtonOk"></param>
        /// <returns></returns>
        private static CustomMaterialMessageBox NormalMessageBox(string text, string title, string textButtonOk)
        {
            CustomMaterialMessageBox messageBox = MessageBox(text, title, textButtonOk);
            messageBox.BtnCancel.Visibility = Visibility.Collapsed;
            return messageBox;
        }

        /// <summary>
        /// créé la popup personnaliser pour l'affichage de message
        /// </summary>
        /// <param name="text"></param>
        /// <param name="title"></param>
        /// <param name="textButtonOk"></param>
        /// <returns></returns>
        private static CustomMaterialMessageBox MessageBox(string text, string title, string textButtonOk)
        {
            return new CustomMaterialMessageBox
            {
                TxtMessage = { Text = text, Foreground = Brushes.Black },
                TxtTitle = { Text = title, Foreground = Brushes.White },
                BtnOk = { Content = textButtonOk, Background = Brushes.Red, BorderBrush = Brushes.DarkRed },
                BtnCancel = { Content = "Cancel", Background = Brushes.DarkRed, BorderBrush = Brushes.DarkRed },
                MainContentControl = { Background = Brushes.Transparent },
                TitleBackgroundPanel = { Background = Brushes.DarkRed },

                BorderBrush = Brushes.DarkRed
            };
        }


    }
}
