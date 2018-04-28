using System.Windows.Forms;
using System;

namespace Tools
{
    public static class FilePopupOpener
    {
        

        /// <summary>
        /// ouvre un dialogue de saisi de fichier
        /// </summary>
        /// <returns></returns>
        public static string OpenExcelFileSelector()
        {
            string path = String.Empty;
            OpenFileDialog fileChooser = new OpenFileDialog();
            if (fileChooser.ShowDialog() == DialogResult.OK)
            {
                path = fileChooser.FileName;
            }
            return path;
        }
       
    }
}
