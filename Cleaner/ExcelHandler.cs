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

        /// <summary>
        /// This method only create an instance of the FileInfo and associate to an ExcelPackage
        /// </summary>
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
            ExcelWorksheet workSheet = pck.Workbook.Worksheets[1];
            var start = workSheet.Dimension.Start;
            var end = workSheet.Dimension.End;
            for (int i = start.Column+1;i <= end.Column;i++)
            {
                for (int j = start.Row;j <= end.Row;j++)
                {
                    string strValue = workSheet.Cells[i,j].Value == null ? string.Empty : workSheet.Cells[i,j].Value.ToString();
                    //MessageBox.Show(strValue);
                }
            }
        }

        public void editColors()
        {
            //Start with the worksheet in the position 1 not 0 
            ExcelWorksheet workSheet = pck.Workbook.Worksheets[1];
            var start = workSheet.Dimension.Start;
            var end = workSheet.Dimension.End;
            BD mybd = new BD();

            mybd.OpenConnection();
            for (int i = start.Row + 1; i <= end.Row; i++)
            {
                if (workSheet.Cells[i, start.Column].Value != null)
                {
                    string codigo = workSheet.Cells[i, 5].Value == null ? "No especificado" : workSheet.Cells[i, 5].Value.ToString();
                    string colores = workSheet.Cells[i, 9].Value == null ? "No especificado" : workSheet.Cells[i, 9].Value.ToString();
                    int inserted = mybd.Colors(colores,codigo);
                }
            }
        }

        public void Edit()
        {
            //Start with the worksheet in the position 1 not 0 
            ExcelWorksheet workSheet = pck.Workbook.Worksheets[1];
            BD mybd = new BD();

            var start = workSheet.Dimension.Start;
            var end = workSheet.Dimension.End;
            
            string today = DateTime.Now.ToString("MM/dd/yyyy");
            string []temps = today.Split('/');
            string stoday = temps[2] + "-" + temps[0] + "-" + temps[1] ;
            string query = "insert into Articulos(Nombre,Descripcion,Codigo,Medidas,Material,Precio,PrecioPublico,IdCategoria,IdSubCategoria,FechaAlta,IdProveedor) values('{0}','{1}','{2}','{3}','{4}','{5}','{6}',{7},{8},'{9}',{10})";
            mybd.OpenConnection();

            string query2 = "select count(*) IdCategoria from Categorias where NombreCategoria = '{0}'";
            for (int i = start.Row + 1; i <= end.Row; i++)
            {
                if(workSheet.Cells[i, start.Column].Value != null)
                {
                    string categoria = workSheet.Cells[i, 2].Value == null ? "No especificado" : workSheet.Cells[i, 2].Value.ToString();
                    string subcategoria = workSheet.Cells[i, 3].Value == null ? "No especificado" : workSheet.Cells[i, 3].Value.ToString();
                    string nombre = workSheet.Cells[i, 4].Value == null ? "No especificado" : workSheet.Cells[i, 4].Value.ToString();
                    string codigo = workSheet.Cells[i, 5].Value == null ? "No especificado" : workSheet.Cells[i, 5].Value.ToString();
                    string descripcion = workSheet.Cells[i, 6].Value == null ? "No especificado" : workSheet.Cells[i, 6].Value.ToString();
                    string medidas = workSheet.Cells[i, 7].Value == null ? "No especificado" : workSheet.Cells[i, 7].Value.ToString();
                    string material = workSheet.Cells[i, 8].Value == null ? "No especificado" : workSheet.Cells[i, 8].Value.ToString();
                    string colores = workSheet.Cells[i, 9].Value == null ? "No especificado" : workSheet.Cells[i, 9].Value.ToString();
                    string precio = workSheet.Cells[i, 10].Value == null ? "No especificado" : workSheet.Cells[i, 10].Value.ToString();
                    string precio_publico = workSheet.Cells[i, 11].Value == null ? "No especificado" : workSheet.Cells[i, 11].Value.ToString();

                    query2 = String.Format(query2,categoria);
                    string idcat = mybd.GetData(query2,categoria);
                    string idsubcat = String.Empty;
                    if (subcategoria != "No especificado")
                    {
                        query2 = "select count(*) IdCategoria from Categorias where NombreCategoria='{0}' and IdPapa!='0';";
                        query2 = String.Format(query2, categoria);
                        idsubcat = mybd.GetData(query2, categoria); 
                    } else
                        idsubcat = "0";

                    int success = mybd.Fill(query,categoria,subcategoria,nombre,codigo,descripcion,medidas,material,colores,precio,precio_publico,idcat,idsubcat,stoday); 
                }
            }
            mybd.CloseConnection();
        }

        /// <summary>
        /// This method only open the Excel application with the file
        /// </summary>
        public void openExcel() {
            System.Diagnostics.Process.Start(file_name);
        }

    }
}
