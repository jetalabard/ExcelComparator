

using System;
using System.Globalization;

namespace Comparator
{
    public class CompareDate
    {
        internal void ManageDateTime(object cell1, object cell2, int i, int columnIndexFile1, int columnIndexFile2, string CurrentColumn)
        {
            DateTime date1;
            DateTime date2;
            try
            {
                date1 = (DateTime)cell1;
                if (cell2 is DateTime)
                {
                    VerifyDate(cell2, i, columnIndexFile1, columnIndexFile2, date1, CurrentColumn);
                }
                else
                {
                    DateTime dDate;
                    if (DateTime.TryParse((string)cell1, out dDate))
                    {
                        String.Format("{0:d/MM/yyyy}", dDate);
                        CompareDateHour(date1, i, columnIndexFile1, columnIndexFile2, dDate, CurrentColumn);
                    }
                    else
                    {
                        ManageMoreComplexDate(cell2, i, columnIndexFile1, columnIndexFile2, date1, CurrentColumn);
                    }
                }
            }
            catch (Exception)
            {
                date2 = (DateTime)cell2;

                if (cell1 is DateTime)
                {
                    VerifyDate(cell1, i, columnIndexFile1, columnIndexFile2, date2,CurrentColumn);
                }
                else
                {
                    DateTime dDate;
                    if (DateTime.TryParse((string)cell1, out dDate))
                    {
                        String.Format("{0:d/MM/yyyy}", dDate);
                        CompareDateHour(date2, i, columnIndexFile1, columnIndexFile2, dDate, CurrentColumn);
                    }
                    else
                    {
                        ManageMoreComplexDate(cell1, i, columnIndexFile1, columnIndexFile2, date2, CurrentColumn);
                    }
                }
            }
        }

        private void ManageMoreComplexDate(object cell2, int i, int columnIndexFile1, int columnIndexFile2, DateTime date1, string CurrentColumn)
        {
            if (cell2.Equals("NULL"))
            {
                if (date1.Year >= 1950)
                {
                    MessageInitializer.Init(i, columnIndexFile1, columnIndexFile2, CurrentColumn);
                }
            }
            else
            {
                DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;

                int weekNum = dfi.Calendar.GetWeekOfYear(
                         date1,
                         CalendarWeekRule.FirstDay,
                         DayOfWeek.Monday);
                string[] splitter = cell2.ToString().Split('/');
                string week = splitter[0];
                string year = splitter[1];
                if (!Convert.ToString(weekNum).Equals(week) || !date1.Year.Equals(year))
                {
                    MessageInitializer.Init(i, columnIndexFile1, columnIndexFile2, CurrentColumn);
                }
            }
        }

        private void VerifyDate(object cell2, int i, int columnIndexFile1, int columnIndexFile2, DateTime date1, String CurrentColumn)
        {
            DateTime date2 = (DateTime)cell2;
            CompareDateHour(date1, i, columnIndexFile1, columnIndexFile2, date2, CurrentColumn);
        }


        private void CompareDateHour(DateTime date1, int i, int columnIndexFile1, int columnIndexFile2, DateTime date2, string CurrentColumn)
        {
            if (!date1.Day.Equals(date2.Day))
            {
                MessageInitializer.Init(i, columnIndexFile1, columnIndexFile2, CurrentColumn);
            }
            else if (!date1.Month.Equals(date2.Month))
            {
                MessageInitializer.Init(i, columnIndexFile1, columnIndexFile2, CurrentColumn);
            }
            else if (!date1.Year.Equals(date2.Year))
            {
                MessageInitializer.Init(i, columnIndexFile1, columnIndexFile2, CurrentColumn);
            }
            else if (!date1.Hour.Equals(date2.Hour))
            {
                MessageInitializer.Init(i, columnIndexFile1, columnIndexFile2, CurrentColumn);
            }
            else if (!date1.Minute.Equals(date2.Minute))
            {
                MessageInitializer.Init(i, columnIndexFile1, columnIndexFile2, CurrentColumn);
            }

        }

    }
}
