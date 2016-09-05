using System;
using IExcel = Microsoft.Office.Interop.Excel;

namespace SQLExecution
{
    public interface IResultObject : IDisposable
    {
        ISqlCommandExecution SqlExecution { get; }
        
        void WriteLog(string message);

        void Execute();
    }
}
