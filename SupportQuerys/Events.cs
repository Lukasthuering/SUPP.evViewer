using System;
using System.IO;
using System.Linq;
using System.Text;
using BiCA.Sabas.Extension.V2;
using OfficeOpenXml;

namespace SupportQuerys
{
    class Events : ManagedScript
    {
        protected override void OnInitializing()
        {
            
        }

        protected override void OnExecuting()
        {
            #region read
            //Read eventlog

            StringBuilder sb = new StringBuilder();
            System.Diagnostics.EventLog log = new
            System.Diagnostics.EventLog("Application");



            DateTime dt = DateTime.MinValue;
            Console.WriteLine(dt.ToString());


            var package = new ExcelPackage(new FileInfo(@"C:\Users\CH-TLU\Desktop\Eventlog" + Runtime.System.Street + "_" + Runtime.System.PostalCode + "_" + DateTime.Now.ToString("yyyyMMdd_hhmmss") + ".xlsx"));
            package.Workbook.Worksheets.Add("Events");
            ExcelWorksheet sheet = package.Workbook.Worksheets.First();

            int row = 1;

            #endregion


            #region write & save
            //Write eventlog to excel file

            foreach (System.Diagnostics.EventLogEntry entry in log.Entries)
            {
                if (entry.TimeGenerated > dt)
                {

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

            sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
            package.Save();
        }

            #endregion

        protected override void OnFinalizing()
        {
     
        }

        protected override void OnErrorOccurred(Exception exception)
        {
            
        }
    }
}
