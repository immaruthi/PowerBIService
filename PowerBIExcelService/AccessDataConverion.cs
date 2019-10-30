/*
 * Maruthi Pallamalli - Pactera Technologies
 * 
 */ 
using Microsoft.Office.Interop.Access.Dao;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Access;
using Microsoft.Vbe.Interop;
using Application = Microsoft.Office.Interop.Access.Application;
using System.IO;
using System.Threading;
using System.Data.OleDb;
using System.Data.Sql;
using System.Data.SqlClient;
using PowerBIExcelService.DataModels;

namespace PowerBIExcelService
{
    public partial class AccessDataConverion : ServiceBase
    {

        private Timer Schedular;
        int dueTime;
        private volatile bool _requestStop = false;

        public AccessDataConverion()
        {
            InitializeComponent();
            
        }

        // Export Query as Excel (*.xlsx supported only)
        private static void ExportQuery(string databaseLocation, string queryNameToExport, string locationToExportTo)
        {
            var application = new Application();
            application.OpenCurrentDatabase(databaseLocation);
            application.DoCmd.TransferSpreadsheet(AcDataTransferType.acExport, AcSpreadSheetType.acSpreadsheetTypeExcel12Xml,
                                                  queryNameToExport, locationToExportTo, true);

            //application.DoCmd.CopyObject(null, null, AcObjectType.acTable, "RawDataReport");

            application.CloseCurrentDatabase();
            application.Quit();
            Marshal.ReleaseComObject(application);
        }

        private static void DirectoryCreation(string[] recommendedInputPaths)
        {
            try
            {
                using (EventLog eventLog = new EventLog())
                {
                    eventLog.Source = "Application";
                    foreach (string recommendedPath in recommendedInputPaths)
                    {
                        if (!System.IO.Directory.Exists(recommendedPath))
                        {
                            eventLog.WriteEntry("Directories Initialization " + recommendedPath);
                            System.IO.Directory.CreateDirectory(recommendedPath);
                            eventLog.WriteEntry("Directories Initialization " + recommendedPath);
                        }
                    }
                }
            }
            catch(Exception exeption)
            {
                using (EventLog eventLog = new EventLog())
                {
                    eventLog.Source = "Application";
                    eventLog.WriteEntry(exeption.Message, EventLogEntryType.Error);
                }
            }
        }

        protected override void OnStart(string[] args)
        {
            dueTime = int.Parse(System.Configuration.ConfigurationSettings.AppSettings["IntervalMinutes"]) * 60000;
            Schedular = new Timer(new TimerCallback(ScheduleTasksCallBack));
            ScheduleTasksCallBack(null);
        }

        private void ScheduleTasksCallBack(object e)
        {
            try
            {

                if (_requestStop)
                {
                    return;
                }

                string[] recommendedInputPaths = new string[] { @"D:\Transformations\Input", @"D:\Transformations\OutPut", @"D:\Transformations\Processed" };
                DirectoryCreation(recommendedInputPaths);

                // Automate RawDataReport Query for Incoming Ms-Access files (*.accdb supported only)
                #region Microsoft.Office.Interop.Access.Dao(version 16.0)

                string[] accdbDirectory = System.IO.Directory.GetFiles(recommendedInputPaths[0], "*.accdb");

                Directory.SetCurrentDirectory(AppDomain.CurrentDomain.BaseDirectory);
                //String Root = Directory.GetCurrentDirectory();

                string rawDataReportQuery = System.IO.File.ReadAllText(Path.Combine(Directory.GetCurrentDirectory(), @"DBObjs\RawDataReport.txt"));

                foreach (string accdbFile in accdbDirectory)
                {
                    try
                    {
                        string accdbFileName = Path.GetFileName(accdbFile);
                        DBEngine dBEngine = new DBEngine();
                        var openDb = dBEngine.OpenDatabase(accdbFile);
                        Console.WriteLine("Preparing Query in desired database");
                        openDb.CreateQueryDef("RawDataReports", rawDataReportQuery);
                        Console.WriteLine("Query Preparation Done");

                        //Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\NKG-TH-CADEMS-BE-Vasari-LP2-PPPL-Ambient-v1.accdb
                        string connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " + accdbFile + ";";
                        DataTable resultsDataset = new DataTable();
                        using (OleDbConnection conn = new OleDbConnection(connString))
                        {
                            OleDbCommand cmd = new OleDbCommand("SELECT * FROM RawDataReports", conn);
                            conn.Open();
                            OleDbDataAdapter adapter = new OleDbDataAdapter(cmd);
                            adapter.Fill(resultsDataset);
                        }

                        IEnumerable<ProgramFilters> programFiltersList =
                            resultsDataset.AsEnumerable()
                            .Select(
                                row =>
                                new
                                {
                                    TestName = row.Field<string>("Test name"),
                                    ProjectPhase = row.Field<string>("Project phase"),
                                    ProgramSKU = row.Field<string>("Program & SKU"),
                                    TestCondition = row.Field<string>("Test Condition")
                                })
                                .Distinct().Select(x => new ProgramFilters()
                                {
                                    ProgramSKU = x.ProgramSKU,
                                    ProjectPhase = x.ProjectPhase,
                                    TestCondition = x.TestCondition,
                                    TestName = x.TestName
                                });

                        List<ServiceDataModel> serviceDataModel = new List<ServiceDataModel>();

                        foreach(ProgramFilters programFiltersForData in programFiltersList)
                        {
                            //DataRow[] dataRows = 
                            //    resultsDataset.Select("Test name = "+ programFiltersForData .TestName+ " AND (Project phase = "+ programFiltersForData.ProjectPhase+ " AND Program & SKU = "+ programFiltersForData.ProgramSKU + " and Test Condition ="+ programFiltersForData.TestCondition + ")");

                            var result = 
                                resultsDataset.AsEnumerable()
                                .Where(
                                    x => x.Field<string>("Test name").Contains(programFiltersForData.TestName) &&
                                         x.Field<string>("Project phase").Contains(programFiltersForData.ProjectPhase) &&
                                         x.Field<string>("Program & SKU").Contains(programFiltersForData.ProgramSKU) &&
                                         x.Field<string>("Test Condition").Contains(programFiltersForData.TestCondition)).CopyToDataTable();

                            serviceDataModel.Add(new ServiceDataModel()
                            {
                                dataTables = result,
                                programFilters = programFiltersForData
                            });      
                        }



                        using (SqlConnection connection = new SqlConnection(System.Configuration.ConfigurationSettings.AppSettings["CentralizedDB"]))
                        {
                            connection.Open();
                            foreach (ServiceDataModel serviceData in serviceDataModel)
                            {

                                string prepareCheckStatement = 
                                    "select * from RawDataReport where Test_name = '"+ serviceData.programFilters.TestName+ "' and Project_phase='" +serviceData.programFilters.ProjectPhase + "' and Program_SKU='" + serviceData.programFilters.ProgramSKU + "' and Test_Condition = '" + serviceData.programFilters.TestCondition + "'";
                                SqlCommand sqlCommand = new SqlCommand(prepareCheckStatement, connection);
                                SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCommand);
                                DataSet dataSet = new DataSet();
                                sqlDataAdapter.Fill(dataSet);

                                if(dataSet.Tables[0].Rows.Count>0)
                                {
                                    // Need to update this DataSet
                                    string deleteStatement = "Delete From RawDataReport where Test_name = '" + serviceData.programFilters.TestName + "' and Project_phase='" + serviceData.programFilters.ProjectPhase + "' and Program_SKU='" + serviceData.programFilters.ProgramSKU + "' and Test_Condition = '" + serviceData.programFilters.TestCondition + "'";
                                    sqlCommand = new SqlCommand(deleteStatement, connection);
                                    sqlCommand.ExecuteNonQuery();
                                    BulkCopyForDataTable(connection, serviceData);
                                }
                                else
                                {
                                    BulkCopyForDataTable(connection, serviceData);
                                }

                            }
                        }


                        string fileSavePath = recommendedInputPaths[1] + "\\" + Path.GetFileNameWithoutExtension(accdbFileName) + DateTime.Now.ToString("yyyy-dd-M--HH-mm-ss") + ".xlsx";
                        Console.WriteLine("Exporting as an Excel File");
                        ExportQuery(accdbFile, "RawDataReports", fileSavePath);
                        Console.WriteLine("Your Excel file is ready at" + fileSavePath);
                        openDb.DeleteQueryDef("RawDataReports");
                        Console.WriteLine("Disposing objects and Restoring to previous version ! original version");
                        openDb.Close();
                        string moveDestinationPath = Path.Combine(recommendedInputPaths[2], Path.GetFileName(accdbFileName));

                        File.Move(accdbFile, moveDestinationPath + DateTime.Now.ToString("_MMMdd_yyyy_HHmmss") + ".accdb");
                    }
                    catch (Exception exception)
                    {
                        using (EventLog eventLog = new EventLog())
                        {
                            eventLog.Source = "Application";
                            eventLog.WriteEntry(exception.Message, EventLogEntryType.Error);
                        }

                        if (exception.Message.Contains("Object 'RawDataReports' already exists."))
                        {
                            DBEngine dBEngine = new DBEngine();
                            var openDb = dBEngine.OpenDatabase(accdbFile);
                            openDb.DeleteQueryDef("RawDataReports");
                            openDb.Close();
                        }
                    }
                }

                #endregion
            }
            catch (Exception exception)
            {
                using (EventLog eventLog = new EventLog())
                {
                    eventLog.Source = "Application";
                    eventLog.WriteEntry(exception.Message, EventLogEntryType.Error);
                }
            }

            Schedular.Change(dueTime, Timeout.Infinite);
        }

        private static void BulkCopyForDataTable(SqlConnection connection, ServiceDataModel serviceData)
        {
            using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(connection))
            {
                sqlBulkCopy.DestinationTableName = "RawDataReport";

                try
                {
                    sqlBulkCopy.WriteToServer(serviceData.dataTables);
                }
                catch (Exception ex)
                {

                }
            }
        }

        protected override void OnStop()
        {
            _requestStop = true;
        }
    }
}
