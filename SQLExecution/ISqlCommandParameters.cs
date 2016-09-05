using System;

namespace SQLExecution
{
    public interface ISqlCommandParameters
    {

        String StoredProcedure { get; set;}
        String SqlServerInstance { get; set; }
        String DatabaseName { get; set; }
        Int32 CommandTimeOut { get; set; }
        Boolean UseIntegratedSecurity { get; set; }
        String UserId { get; set; }
        String Password { get; set; }
    }
}
