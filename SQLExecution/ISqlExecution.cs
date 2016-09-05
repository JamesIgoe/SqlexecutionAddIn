using System;

namespace SQLExecution
{
    public interface ISqlExecution : IDisposable
    {
        void Run();
    }
}
