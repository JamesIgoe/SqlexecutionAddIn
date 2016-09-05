using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using IExcel = Microsoft.Office.Interop.Excel;
using System.Threading;

namespace SQLExecution
{
    public delegate void UpdateCallingClass(String message);

    public class SqlExecution : ISqlExecution, IDisposable
    {
        public SqlExecution()
        {
            //
        }
     
        private bool disposed = false;
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~SqlExecution()
        {
            Dispose(false);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                foreach (IResultObject result in Commands)
                {
                    result.Dispose();
                }
            }

            sendMessage = null;

            disposed = true;
        }

   
        # region messaging
        
        private UpdateCallingClass sendMessage;
        public void SubscribeToDelegate(UpdateCallingClass method)
        {
            sendMessage += method;
        }

        public void UnsubscribeFromDelegate(UpdateCallingClass method)
        {
            sendMessage -= method;
        }

        public void WriteBack(string message)
        {
            if (sendMessage != null)
            {
                sendMessage(message);
            }
        }

        #endregion 
        
        # region Code for building list of commands to execute

        public IList<IResultObject> Commands = new List<IResultObject>();
        
        public virtual void Clear()
        {
            foreach (IResultObject result in this.Commands)
            {
                result.Dispose();
            }
            this.Commands.Clear();
        }

        public void AddToList(ISqlCommandParameters parameters)
        {
            this.Commands.Add(new ResultObject(parameters, this.WriteBack));
            WriteBack(String.Format("Command: {0}", parameters.StoredProcedure));
        }

        # endregion

        # region Code for executing SP's threaded

        public virtual void Run()
        { 
            RunThreads();
        }

        public void RunThreads()
        {
            Thread[] threads = new Thread[this.Commands.Count];
            Int32 counter = 0;

            //for each command in commands, start threaded
            //run threaded, until all done or timeout
            foreach (IResultObject command in this.Commands)
            {
                ThreadStart execution = new ThreadStart(command.Execute);
                Thread query = new Thread(execution);

                threads[counter++] = query;

                query.IsBackground = true;
                query.Start();
            }

            for (int threadCounter = 0; threadCounter < this.Commands.Count; threadCounter++)
            {
                threads[threadCounter].Join();
            }
        }

        # endregion

    }
}
