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
                DatabaseName = "SQL2008_600471_codedotnet",
                SqlServerInstance = "sql2k802.discountasp.net",
                StoredProcedure = "SELECT [MonthNum],[MonthText] FROM [SQL2008_600471_codedotnet].[dbo].[Z_tblMonth]",
                UseIntegratedSecurity = false,
                UserId = "SQL2008_600471_codedotnet_user",
                Password = "pauline"
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
                DatabaseName = "SQL2008_600471_codedotnet",
                SqlServerInstance = "sql2k802.discountasp.net",
                StoredProcedure = "SELECT [MonthNum],[MonthText] FROM [SQL2008_600471_codedotnet].[dbo].[Z_tblMonth]",
                UseIntegratedSecurity = false,
                UserId = "SQL2008_600471_codedotnet_user",
                Password = "pauline"
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
                DatabaseName = "SQL2008_600471_codedotnet",
                SqlServerInstance = "sql2k802.discountasp.net",
                StoredProcedure = "SELECT TOP 10 * FROM [SQL2008_600471_codedotnet].[dbo].[Z_tblMonth]",
                UseIntegratedSecurity = false,
                UserId = "SQL2008_600471_codedotnet_user",
                Password = "pauline"
            };

            SQLExecution.SqlCommandParameters param2 = new SqlCommandParameters()
            {
                CommandTimeOut = 60,
                DatabaseName = "SQL2008_600471_codedotnet",
                SqlServerInstance = "sql2k802.discountasp.net",
                StoredProcedure = "SELECT TOP 6 * FROM [SQL2008_600471_codedotnet].[dbo].[Z_tblMonth]",
                UseIntegratedSecurity = false,
                UserId = "SQL2008_600471_codedotnet_user",
                Password = "pauline"
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
