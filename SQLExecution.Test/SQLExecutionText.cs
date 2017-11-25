using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SQLExecution;
using System.Data;
using System.Data.SqlClient;

namespace SQLExecution.Test
{
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

            cmd.Dispose();
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

            using (SQLExecution.SqlExecution exec = new SqlExecution())
            {
                exec.AddToList(param);
                exec.Run();

                using (System.Data.DataSet set = exec.Commands[0].SqlExecution.Data)
                {
                    Assert.IsNotNull(set);
                    Assert.IsTrue(set.Tables[0].Rows.Count > 0);
                }
            }

        }
        
        [TestMethod]
        public void SQLExecutionTwoThreadTest()
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

            SQLExecution.SqlCommandParameters param2 = new SqlCommandParameters()
            {
                CommandTimeOut = 60,
                DatabaseName = "",
                SqlServerInstance = "",
                StoredProcedure = "",
                UseIntegratedSecurity = false,
                UserId = "",
                Password = ""
            };

            using (SQLExecution.SqlExecution exec = new SqlExecution())
            {
                exec.AddToList(param);
                exec.AddToList(param2);

                exec.Run();

                using (System.Data.DataSet set = exec.Commands[0].SqlExecution.Data)
                {
                    Assert.IsTrue(set.Tables[0].Rows.Count == 10);
                }

                using (System.Data.DataSet set = exec.Commands[1].SqlExecution.Data)
                {
                    Assert.IsTrue(set.Tables[0].Rows.Count == 6);
                }
            }
        }
    }
}
