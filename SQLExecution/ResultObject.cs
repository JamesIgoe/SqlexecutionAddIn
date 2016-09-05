using System;
using IExcel = Microsoft.Office.Interop.Excel;
using System.Data;

namespace SQLExecution
{
    public class ResultObject : SQLExecution.IResultObject, IDisposable
    {
        protected ISqlCommandExecution sqlCommandExecution;
        public ISqlCommandExecution SqlExecution
        {
            get { return sqlCommandExecution; }
        }

        private UpdateCallingClass callingClass;
        public ResultObject(ISqlCommandParameters parameter, UpdateCallingClass logMethod)
        {
            sqlCommandExecution = new SqlCommandExecution(parameter);
            callingClass = logMethod;
        }

        private bool disposed = false;
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~ResultObject()
        {
            Dispose(false);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                sqlCommandExecution.Dispose();
            }

            callingClass = null;

            disposed = true;
        }

        public void Execute()
        {
            sqlCommandExecution.Execute();
        }

        public void WriteLog(string message)
        {
            if (callingClass != null)
            {
                callingClass(message);
            }
        }
    }
}
