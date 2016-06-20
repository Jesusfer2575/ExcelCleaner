using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Cleaner
{
    class ExcelHandler
    {
        private string file_name;
        private ExcelPackage pck;
        private FileInfo newFile;

        public ExcelHandler() {}

        public ExcelHandler(string f) {
            this.file_name = f;
        }

        public void open()
        {
            try
            {
                newFile = new FileInfo(file_name);
                pck = new ExcelPackage(newFile);
            }
            catch(Exception ex){
                Console.WriteLine(ex.ToString());
            }
            
        }
        public void addWorkbook() {
            //Add the Content sheet
            var ws = pck.Workbook.Worksheets.Add("Content");
            ws.View.ShowGridLines = false;

            ws.Column(4).OutlineLevel = 1;
            ws.Column(4).Collapsed = true;
            ws.Column(5).OutlineLevel = 1;
            ws.Column(5).Collapsed = true;
            ws.OutLineSummaryRight = true;

            //Headers
            ws.Cells["B1"].Value = "Name";
            ws.Cells["C1"].Value = "Size";
            ws.Cells["D1"].Value = "Created";
            ws.Cells["E1"].Value = "Last modified";
            //ws.Cells["B1:E1"].Style.Font.Bold = true;

            pck.Save();
        }
        
        /// <summary>
        /// This method reads every cell one by one
        /// </summary>
        public void readFile()
        {
            //Start with the worksheet in the position 1 not 0 
            ExcelWorksheet workSheet = pck.Workbook.Worksheets[4];
            var start = workSheet.Dimension.Start;
            var end = workSheet.Dimension.End;
            for (int i = start.Column+1;i <= end.Column;i++)
            {
                for (int j = start.Row;j <= end.Row;j++)
                {
                    string strValue = workSheet.Cells[i,j].Value == null ? string.Empty : workSheet.Cells[i,j].Value.ToString();
                }
            }
        }   

        public void openExcel() {
            System.Diagnostics.Process.Start(file_name);
        }

    }
}
