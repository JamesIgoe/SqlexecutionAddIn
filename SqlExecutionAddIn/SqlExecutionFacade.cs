using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using IExcel = Microsoft.Office.Interop.Excel;

namespace SQLExecutionAddIn
{
    /// <summary>
    /// Interace to expose VSTO/COM obects for COM and Excel
    /// Used by class below
    /// </summary>
    [ComVisible(true)]
    [Guid("9D23A5CA-D036-4895-AE7F-CDE1EB6ECFA3")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ISqlExecutionForVBA
    {
        //signatures for methods acessible to VBA go here
        //exists here and in class
        void SetQprParameters(String currentMonthEnd,
                    String priorMonthEnd,
                    String dataSource,
                    String portfolioCurrent,
                    String portfolioPrior,
                    String portfolioBenchmarkCurrent,
                    String portfolioBenchmarkPrior,
                    String portfolioBenchmarkCurrentWeights,
                    String portfolioBenchmarkPriorWeights,
                    String ratingAgency,
                    String sqlType,
                    String targetWorkbookName,
                    String serverInstance,
                    String databaseName,
                    Int32 commandTimeout,
                    String fullLogPath
                    );

        void ExecuteQpr();

        void AddSPToCollection(String spWithParameters,
                            IExcel.Worksheet targetWorksheet,
                            String targetCellAsString,
                            Boolean includeFieldHeadings,
                            String serverInstance,
                            String databaseName,
                            Int32 commandTimeout
                            );

        void ExecuteSPCollection();
        void ClearSPCollection();
    }

    /// <summary>
    /// Class to expose items to VBA and other COM applications
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class SqlExecutionForVBA : SQLExecutionAddIn.ISqlExecutionForVBA
    {
        # region properties and code for QPR

        SQLExecution.PortfolioParameters parameters = new SQLExecution.PortfolioParameters();
        SQLExecution.SqlExecutionForQPR sql;
        String logPath;

        public SqlExecutionForVBA()
        {

        }

        //methods to expose go here
        //exists here and in interface
        public void SetQprParameters(String currentMonthEnd, 
                            String priorMonthEnd, 
                            String dataSource, 
                            String portfolioCurrent, 
                            String portfolioPrior,
                            String portfolioBenchmarkCurrent,
                            String portfolioBenchmarkCurrentWeights,
                            String portfolioBenchmarkPrior,
                            String portfolioBenchmarkPriorWeights,
                            String ratingAgency,
                            String sqlType,
                            String targetWorkbookName,
                            String serverInstance,
                            String databaseName,
                            int commandTimeout,
                            String fullLogPath
                            )
        {
            parameters.CurrentMonthEnd = Convert.ToDateTime(currentMonthEnd);
            parameters.PriorMonthEnd = Convert.ToDateTime(priorMonthEnd);
            parameters.DataSource = (SQLExecution.CharacteristicDataSource)Enum.Parse(typeof(SQLExecution.CharacteristicDataSource), dataSource);
            parameters.PortfolioCurrent = portfolioCurrent;
            parameters.PortfolioPrior = portfolioPrior;
            parameters.PortfolioBenchmarkCurrent = portfolioBenchmarkCurrent;
            parameters.PortfolioBenchmarkCurrentWeights = portfolioBenchmarkCurrentWeights;
            parameters.PortfolioBenchmarkPrior = portfolioBenchmarkPrior;
            parameters.PortfolioBenchmarkPriorWeights = portfolioBenchmarkPriorWeights;
            parameters.Agency = (SQLExecution.RatingAgency)Enum.Parse(typeof(SQLExecution.RatingAgency), ratingAgency);
            parameters.SqlType = sqlType.Trim().Length > 0 ? (SQLExecution.SPAppend)Enum.Parse(typeof(SQLExecution.SPAppend), sqlType) : SQLExecution.SPAppend.None;
            parameters.TargetWorkbookName = targetWorkbookName;
            parameters.ServerInstance = serverInstance;
            parameters.DatabaseName = databaseName;
            parameters.CommandTimeout = commandTimeout;
            parameters.FullLogPath = fullLogPath;

            logPath = fullLogPath;

            sql = new SQLExecution.SqlExecutionForQPR(parameters);
            sql.SubscribeToDelegate(this.AppendToTextFile);
        }

        public void ExecuteQpr()
        {
            sql.Run();
            sql.UnsubscribeFromDelegate(this.AppendToTextFile);
        }

        private void AppendToTextFile(String lineToAdd)
        {
            if (System.IO.File.Exists(logPath))
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(logPath, true))
                {
                    String fullLineText = String.Format("{0}: {1}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), lineToAdd);
                    file.WriteLine(fullLineText);
                }
            }
        }
        # endregion


        # region Code for generic SQL execution via VBA

        SQLExecution.SqlExecutionForExcel results = new SQLExecution.SqlExecutionForExcel();
        public void AddSPToCollection(String spWithParameters,
                                                IExcel.Worksheet targetWorksheet,
                                                String targetCellAsString,
                                                Boolean includeFieldHeadings,
                                                String serverInstance,
                                                String databaseName,
                                                Int32 commandTimeout
                                                )
        {
            SQLExecution.ISqlCommandParameters parameters = new SQLExecution.SqlCommandParameters();
            parameters.StoredProcedure = spWithParameters;
            parameters.SqlServerInstance = serverInstance;
            parameters.DatabaseName = databaseName;
            parameters.CommandTimeOut = commandTimeout;

            IExcel.Range rng = targetWorksheet.get_Range(targetCellAsString, Type.Missing);

            results.AddToList(parameters, rng, includeFieldHeadings);
        }

        public void ExecuteSPCollection()
        {
            if (logPath!=null && logPath.Length > 0)
            {
                if (System.IO.File.Exists(logPath))
                {
                    results.SubscribeToDelegate(this.AppendToTextFile);
                }
            }

            results.Run();
            results.UnsubscribeFromDelegate(this.AppendToTextFile);
        }

        public void ClearSPCollection()
        {
            results.Clear();
        }

        # endregion
    }
}