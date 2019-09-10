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

                foreach (string accdbFile in accdbDirectory)
                {
                    try
                    {
                        string accdbFileName = Path.GetFileName(accdbFile);
                        DBEngine dBEngine = new DBEngine();
                        var openDb = dBEngine.OpenDatabase(accdbFile); 
                        openDb.CreateQueryDef("RawDataReports", "Select * from DataEntryDefect");
                        string fileSavePath = recommendedInputPaths[1] + "\\" + Path.GetFileNameWithoutExtension(accdbFileName) + DateTime.Now.ToString("yyyy-dd-M--HH-mm-ss") + ".xlsx";
                        ExportQuery(accdbFile, "RawDataReports", fileSavePath);
                        openDb.DeleteQueryDef("RawDataReports");
                        openDb.Close();
                        string moveDestinationPath = Path.Combine(recommendedInputPaths[2], Path.GetFileName(accdbFileName));
                        File.Move(accdbFile, moveDestinationPath);
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


        protected override void OnStop()
        {
            _requestStop = true;
        }
    }
}
