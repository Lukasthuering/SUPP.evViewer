using BiCA.Sabas.Extension.V2;
using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SupportQuerys
{
    class Events : ManagedScript
    {
        
        protected override void OnInitializing()
        {

        }
        
        protected override void OnExecuting()
        {
            #region read eventlog
            //Read eventlog

            System.Diagnostics.EventLog log = new
            System.Diagnostics.EventLog("Application");
           
            DateTime dt = new DateTime();
            Console.WriteLine(dt.ToString());

            //Create Excel file with the correct name
            //Create file on Server
            //  var package = new ExcelPackage(new FileInfo(@"\\ch-bicaap02\Support\TEMP\SUPP.evViewer\" + Runtime.System.SisJobNumber + Runtime.System.Street + "_" + Runtime.System.PostalCode + "_" + DateTime.Now.ToString("yyyyMMdd_hhmmss") + ".xlsx"));

            //Create local file (Testing purposes)
            var package = new ExcelPackage(new FileInfo(@"C:\Temp\" + Runtime.System.SisJobNumber + Runtime.System.Street + "_" + Runtime.System.PostalCode + "_" + DateTime.Now.ToString("yyyyMMdd_hhmmss") + ".xlsx"));
            //Add new worksheet with the name "Events"
            package.Workbook.Worksheets.Add("Events");
            //Go to the first (and only) worksheet in the Excel sheet
            ExcelWorksheet sheet = package.Workbook.Worksheets.First();

            //Row needs to be set to 2, because there will be a filter in row 1
            int row = 2;

            #endregion
            
            #region write to excel file & save

            //Write eventlog to excel file

            //foreach loop to write the event to the excel
            foreach (System.Diagnostics.EventLogEntry entry in log.Entries)
            {
                if (entry.TimeGenerated > dt)
                {
                    sheet.SelectedRange["A1"].Value = "";
                    sheet.Cells[row, 1].Value = Runtime.System.SiteServer.HostAddress;
                    sheet.Cells[row, 2].Value = Runtime.System.SystemID;
                    sheet.Cells[row, 3].Value = entry.Source;
                    sheet.Cells[row, 4].Value = entry.EntryType;
                    sheet.Cells[row, 5].Value = entry.TimeGenerated.Date.ToLongDateString();
                    sheet.Cells[row, 6].Value = entry.TimeWritten.ToLongTimeString();
                    sheet.Cells[row, 7].Value = entry.Message;
                    row++;
                }
            }
            //Sets the Column size to automatic for a better view
            sheet.Cells[sheet.Dimension.Address].AutoFitColumns();

            //saves the file
            package.Save();

            #endregion
        }

     

        

        protected override void OnFinalizing()
        {
            
        }

        protected override void OnErrorOccurred(Exception exception)
        {

        }
    }
}
