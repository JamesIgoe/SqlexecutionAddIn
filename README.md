# SqlexecutionAddIn

A library that provides simplified threaded and non-threaded execution of SQL statements. It can be used to expose an interface to Microsoft Excel clients, providing threaded, async operations to a single-thread application. You will need to build this on a computer with the MS Office DLL's and/or PIA's. 


NuGet Package

* The NuGet package is avilable here: https://www.nuget.org/packages/SQLExecution/


SQL Execution Documentation (NuGet Package)

* 1.3.0: Added non-trusted execution parameters, corrected threaded execution
* 1.2.0: Made class virtual to allow instantiation
* 1.1.1: Corrected assembly information
* 1.1.0: Corrected class hierarchy


SQL Execution Tests

* See support the site, http://comparative-advantage.com/code/SQL_ExecutionHelp.php, for full usage syntaz

    [TestClass]
    public class SQLExecutionText
    {
        [TestMethod]
        public void CreateSqlExecutionClasssTest()
        {
            SQLExecution.SqlExecution sqlClass = new SQLExecution.SqlExecution();
            Assert.IsNotNull(sqlClass);
        }

        [TestMethod]
        public void SQLCommandExecutionTest()
        {
            SQLExecution.SqlCommandParameters param = new SqlCommandParameters()
            {
                CommandTimeOut = 60,
                DatabaseName = "",
                SqlServerInstance = "",
                StoredProcedure = "",
                UseIntegratedSecurity = false,
                UserId = "",
                Password = ""
            };

            SQLExecution.SqlCommandExecution cmd = new SqlCommandExecution(param);

            cmd.Execute();

            System.Data.DataSet set = cmd.Data;

            Assert.IsNotNull(set);
            Assert.IsTrue(set.Tables[0].Rows.Count > 0);
        }

        [TestMethod]
        public void SQLExecutionTest()
        {
            SQLExecution.SqlCommandParameters param = new SqlCommandParameters()
            {
                CommandTimeOut = 60,
                DatabaseName = "",
                SqlServerInstance = "",
                StoredProcedure = "",
                UseIntegratedSecurity = false,
                UserId = "",
                Password = ""
            };

            SQLExecution.SqlExecution exec = new SqlExecution();

            exec.AddToList(param);

            exec.Run();

            System.Data.DataSet set = exec.Commands[0].SqlExecution.Data;

            Assert.IsNotNull(set);
            Assert.IsTrue(set.Tables[0].Rows.Count > 0);
        }


VBA Usage

An example using the code to execute SQL asynchronously and write it out to different sheets. In effect, the execution time is nearer to the execution time of the slowest command object, rather than being the sum of execution times. 

Some Notes:

* The data classes used above can be used instead of the ones below if you only want to execute SQL and collect data
The classes below both retrieve data and write it out to sheets
* Some fields in the code below are global parameters as string for server instance, database name and timeout, prefaced by gstr, which can be passed in as variables instead
* Each item will execute independently, and do not need to be pointing at the same server/database
* See support the site, http://comparative-advantage.com/code/SQL_ExecutionHelp.php, for usage syntax

