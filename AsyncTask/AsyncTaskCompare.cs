using Comparator;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace AsyncTask
{
    public class AsyncTaskCompare
    {
        /// <summary>
        /// chemin des fichiers à comparer
        /// </summary>
        private readonly List<Worksheet> WorkSheets;
        private readonly Dictionary<int, string> columnsFile1;
        private readonly Dictionary<int, string> columnsFile2;

        public delegate void OnWorkerMethodCompleteDelegate(string result);
        public event OnWorkerMethodCompleteDelegate OnWorkerComplete;

        public AsyncTaskCompare(List<Worksheet> WorkSheets, Dictionary<int, string> columnsFile1, Dictionary<int, string> columnsFile2)
        {
            this.WorkSheets = WorkSheets;
            this.columnsFile1 = columnsFile1;
            this.columnsFile2= columnsFile2;
        }

        /// <summary>
        /// sauvegarde les données du fichier en base, compare avec la BDD et créé un rapport
        /// </summary>
        public void WorkerMethod()
        {
            string result = new CompareFile(WorkSheets).CompareTwoSheet(columnsFile1, columnsFile2);
            OnWorkerComplete(result);
        }

    }
}
