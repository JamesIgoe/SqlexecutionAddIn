using System;
using IExcel = Microsoft.Office.Interop.Excel;
using System.Data;

namespace SQLExecution
{
    public class ResultObjectForExcel : IResultObjectForExcel, IDisposable
    {
        private ISqlCommandExecution sqlExecution;
        public ISqlCommandExecution SqlExecution
        {
            get { return sqlExecution; }
        }

        private Boolean includeHeader;
        public Boolean IncludeHeader
        {
            get { return includeHeader; }
            set { includeHeader = value; }
        }

        private IExcel.Range targetRange;
        public IExcel.Range TargetRange
        {
            get { return targetRange; }
        }

        private UpdateCallingClass callingClass;
        public ResultObjectForExcel(ISqlCommandParameters parameters, IExcel.Range target, Boolean includeFieldHeadings, UpdateCallingClass logMethod)
        {
            targetRange = target;
            includeHeader = includeFieldHeadings;
            sqlExecution = new SqlCommandExecution(parameters);
            callingClass = logMethod;
        }

        private bool disposed = false;
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~ResultObjectForExcel()
        {
            Dispose(false);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                sqlExecution.Dispose();
            }

            callingClass = null;

            ComCleanup.FinalReleaseAndNull(targetRange);
            ComCleanup.GarbageCleanup();

            disposed = true;
        }

        public void Execute()
        {
            sqlExecution.Execute();
        }

        public void WriteLog(string message)
        {
            if (callingClass != null)
            {
                callingClass(message);
            }
        }

        public void WriteData()
        {
            Int32 rowCount;
            Int32 colCount;

            //report error or data empty
            if (sqlExecution.Error != null)
            {
                WriteLog(string.Format("Error: {0}", sqlExecution.Error.Message));
            }
            else if (sqlExecution.Data.Tables.Count == 0)
            {
                WriteLog(string.Format("No tables returned for {0}", sqlExecution.SqlParameters.StoredProcedure));
            }
            else if (sqlExecution.Data.Tables[0].Rows.Count == 0)
            {
                WriteLog(string.Format("No rows returned for {0}", sqlExecution.SqlParameters.StoredProcedure));
            }
            else
            {
                rowCount = sqlExecution.Data.Tables[0].Rows.Count;
                colCount = sqlExecution.Data.Tables[0].Columns.Count;

                //has table, so right headers if wanted
                if (includeHeader == true)
                {
                    for (Int32 cols = 0; cols < colCount; cols++)
                    {
                        IExcel.Range rng = targetRange.get_Offset(0, cols);
                        rng.Value = sqlExecution.Data.Tables[0].Columns[cols].ToString();
                        ComCleanup.FinalReleaseAndNull(rng);
                    }
                }
                
                Int32 rowToStart = includeHeader ? 1 : 0;
                for (Int32 rows = 0; rows < rowCount; rows++)
                {
                    DataRow rw = sqlExecution.Data.Tables[0].Rows[rows];

                    for (Int32 cols = 0; cols < colCount; cols++)
                    {
                        IExcel.Range rng = targetRange.get_Offset(rows + rowToStart, cols);
                        rng.Value = rw[cols].ToString();
                        ComCleanup.FinalReleaseAndNull(rng);
                    }
                }
            }
        }

        public IExcel.Worksheet ActivateWorksheet(IExcel.Range target)
        {
            IExcel.Worksheet wks = (IExcel.Worksheet)target.Worksheet;
            ((IExcel._Worksheet)wks).Activate();
            return wks;
        }
    }
}
