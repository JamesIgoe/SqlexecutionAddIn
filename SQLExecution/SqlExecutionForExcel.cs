using System;
using System.Collections.Generic;
using System.Data;
using System.Threading;
using IExcel = Microsoft.Office.Interop.Excel;

namespace SQLExecution
{
    public class SqlExecutionForExcel : SqlExecution
    {

        public SqlExecutionForExcel()
        {
            //nothing
        }

        private bool disposed = false;
        ////inherited
        //public void Dispose()
        //{
        //    Dispose(true);
        //    GC.SuppressFinalize(this);
        //}

        //~SqlExecutionForExcel()
        //{
        //    Dispose(false);
        //}

        new protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                foreach (IResultObjectForExcel result in Commands)
                {
                    result.Dispose();
                }
            }

            disposed = true;
        }

        # region Code for building list of commands to execute

        new protected IList<ResultObjectForExcel> Commands = new List<ResultObjectForExcel>();

        public override void Clear()
        {
            foreach (IResultObjectForExcel result in this.Commands)
            {
                result.Dispose();
            }
            this.Commands.Clear();
        }

        public void AddToList(ISqlCommandParameters parameters, IExcel.Range target, Boolean includeFieldHeadings)
        {
            this.Commands.Add(new ResultObjectForExcel(parameters, target, includeFieldHeadings, base.WriteBack));
            base.WriteBack(String.Format("Command: {0}", parameters.StoredProcedure));
        }

        # endregion

        # region Code for executing SP's threaded, and writing to Excel

        public override void Run()
        {
            base.RunThreads();
            this.WriteDataToSheets();
        }

        private void WriteDataToSheets()
        {
            foreach (ResultObjectForExcel command in this.Commands)
            {
                if (command.TargetRange != null)
                {
                    command.ActivateWorksheet(command.TargetRange);
                    command.WriteData();
                    command.Dispose();
                    ComCleanup.GarbageCleanup();
                }
            }
        }

        # endregion
    }
}
