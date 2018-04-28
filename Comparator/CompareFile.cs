using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;


namespace Comparator
{
    public class CompareFile
    {
        public string Result { get; private set; }
        public string CurrentColumn { get; private set; }
        private readonly CompareDate DateComparator;
        private readonly List<Worksheet> Worksheets;

        public CompareFile(List<Worksheet> worksheets)
        {
            Worksheets = worksheets;
            DateComparator = new CompareDate();
        }

        public string CompareTwoSheet(Dictionary<int, string> columnsFile1, Dictionary<int, string> columnsFile2)
        {

            if (!SameKeys(columnsFile1, columnsFile2))
            {
                Tools.Message.Show("Les deux fichiers n'ont pas les mêmes colonnes à comparer.");
            }
            else
            {
                int rows = GetNumberRowsInSheet(Worksheets);
                for (int i = 2; i <= rows; i++)
                {
                    foreach (string column in columnsFile1.Values)
                    {
                        CurrentColumn = column;
                        int indexColumFile1 = columnsFile1.Where(col => col.Value.Equals(column)).Select(col => col.Key).First();
                        int indexColumFile2 = columnsFile2.Where(col => col.Value.Equals(column)).Select(col => col.Key).First();
                        object cell1 = (Worksheets[0].Cells[i, indexColumFile1] as Range).Value;
                        object cell2 = (Worksheets[1].Cells[i, indexColumFile2] as Range).Value;
                        if (cell1 == null)
                        {
                            cell1 = String.Empty;
                        }
                        if (cell2 == null)
                        {
                            cell2 = String.Empty;
                        }
                        if (cell1 is DateTime || cell2 is DateTime)
                        {
                            DateComparator.ManageDateTime(cell1, cell2, i, indexColumFile1, indexColumFile2,CurrentColumn);
                        }
                        else
                        {
                            ManageOther(cell1, cell2, i, indexColumFile1, indexColumFile2);
                        }
                    }
                }
            }
            if (string.IsNullOrEmpty(Result))
            {
                Result = "File Equal";
            }
            return Result;
        }

        private bool SameKeys(Dictionary<int, string> one, Dictionary<int, string> two)
        {
            if (one.Count != two.Count)
                return false;
            foreach (var value in one.Values)
            {
                if (!two.ContainsValue(value))
                    return false;
            }
            return true;
        }



        private void ManageOther(object cell1, object cell2, int i, int columnIndexFile1, int columnIndexFile2)
        {
            try
            {
                if (!cell1.Equals(cell2))
                {
                    MessageInitializer.Init(i, columnIndexFile1, columnIndexFile2,CurrentColumn);
                }
            }
            catch (Exception excep)
            {
                Tools.Message.Show(excep.Message);
            }
        }

      
     
        private int GetNumberRowsInSheet(List<Worksheet> worksheets)
        {
            Range xlRange = worksheets[0].UsedRange;
            Range xlRange2 = worksheets[1].UsedRange;
            int rows = 0;
            if (xlRange2.Rows.Count >= xlRange.Rows.Count)
            {
                rows = xlRange.Rows.Count;
            }
            else
            {
                rows = xlRange2.Rows.Count;
            }
            return rows;
        }

    }
}
