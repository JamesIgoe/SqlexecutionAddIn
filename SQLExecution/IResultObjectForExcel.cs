using System;
using IExcel = Microsoft.Office.Interop.Excel;

namespace SQLExecution
{
    public interface IResultObjectForExcel: IResultObject, IDisposable
    {
        bool IncludeHeader { get; set; }
        void WriteData();
        
        IExcel.Range TargetRange { get; }
        IExcel.Worksheet ActivateWorksheet(Microsoft.Office.Interop.Excel.Range target);
    }
}
