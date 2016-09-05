using System;
using System.Data;
using System.Data.SqlClient;

namespace SQLExecution
{
    public class SqlCommandExecution : ISqlCommandExecution, IDisposable
    {
       # region Properties

        String defaultStringSSPI = Properties.Settings.Default.DefaultConnectionStringSSPI;
        String defaultStringNonSSPI = Properties.Settings.Default.DefaultConnectionStringNonSSPI;

        String connectionString;

        SqlConnection connection = null;
        SqlCommand cmd = null;
        SqlDataAdapter adapter = null;

        Exception error;
        public Exception Error
        { 
            get {return error;} 
        }

        private DataSet data = new DataSet();
        public DataSet Data
        {
            get { return data; }
        }

        # endregion

        # region Constructors and Dispose

        ISqlCommandParameters sqlParameters;
        public ISqlCommandParameters SqlParameters
        {
            get { return sqlParameters; }
        }

        private bool disposed = false;
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~SqlCommandExecution()
        {
            Dispose(false);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                data.Dispose();
                adapter.Dispose();
                cmd.Dispose();
                connection.Dispose();
            }

            disposed = true;
        }

        public SqlCommandExecution(ISqlCommandParameters parameters)
        {
            sqlParameters = parameters;
        }

        #endregion

        # region SQL parameter construction and execution

        public void Execute()
        {
            try
            {
                connectionString = GetConnectionString();

                using (connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    using (cmd = GetCommand(connection))
                    { 
                        using (adapter = new SqlDataAdapter())
                        {
                            adapter.SelectCommand = cmd;
                            adapter.Fill(data);
                        }
                    }
                }
            }
            catch (Exception ex)
            { 
                error = ex;
            }
        }

        private SqlCommand GetCommand(SqlConnection connection)
        {
            if (cmd == null)
            {
                cmd = new SqlCommand();
            }
            cmd.CommandText = sqlParameters.StoredProcedure;
            cmd.Connection = connection;
            cmd.CommandType = CommandType.Text;
            cmd.CommandTimeout = sqlParameters.CommandTimeOut;
            return cmd;
        }

        private string GetConnectionString()
        {
            string connectionString = sqlParameters.UseIntegratedSecurity ? defaultStringSSPI : defaultStringNonSSPI;

            connectionString = connectionString.Replace("%ServerAddress%", sqlParameters.SqlServerInstance);
            connectionString = connectionString.Replace("%DataBase%", sqlParameters.DatabaseName);
            connectionString = connectionString.Replace("%UserId%", sqlParameters.UserId);
            connectionString = connectionString.Replace("%Password%", sqlParameters.Password);

            return connectionString;
        }

        # endregion
    }
}
