using System;
using System.Data;

namespace SQLExecution
{
    public interface ISqlCommandExecution: IDisposable
    {
        DataSet Data { get; }
        Exception Error { get; }
        ISqlCommandParameters SqlParameters { get; }
        void Execute();
    }
}
