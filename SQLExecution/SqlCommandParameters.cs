using System;

namespace SQLExecution
{
    public class SqlCommandParameters : ISqlCommandParameters
    {
        public String StoredProcedure { get; set; }
        public String SqlServerInstance { get; set; }
        public String DatabaseName { get; set; }
        public Int32 CommandTimeOut { get; set; }
        public Boolean UseIntegratedSecurity { get; set; }
        public String UserId { get; set; }
        public String Password { get; set; }
    }
}
