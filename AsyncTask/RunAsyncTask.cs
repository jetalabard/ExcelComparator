using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Threading;
using Tools;

namespace AsyncTask
{
    public class RunAsyncTask
    {
        private LoadingWindow popup;

        public Thread Process { get; private set; }

        public void Run(List<Worksheet> WorkSheet, Dictionary<int, string> columnsFile1, Dictionary<int, string> columnsFile2)
        {
            AsyncTaskCompare task = new AsyncTaskCompare(WorkSheet, columnsFile1, columnsFile2);
            task.OnWorkerComplete += new AsyncTaskCompare.OnWorkerMethodCompleteDelegate(OnWorkerMethodComplete);
            ThreadStart tStart = new ThreadStart(task.WorkerMethod);
            Process = new Thread(tStart);
            Process.Start();

            popup = new LoadingWindow(true);
            popup.ShowDialog();
        }

        private void OnWorkerMethodComplete(string result)
        {
            popup.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Normal,
            new System.Action(
            delegate ()
            {
                Message.Show(result);
                popup.Close();
            }
            ));

            if (Process != null && Process.IsAlive)
            {
                Process.Abort();
            }
        }
    }
}
