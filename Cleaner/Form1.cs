using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Cleaner
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string nom = textBox1.Text;
            string name_path = "C:\\Users\\adria\\Google Drive\\CLIENTES\\DOBLECERO\\FUENTE\\DobleCero\\CATALOGO.xlsx";
            ExcelHandler libro = new ExcelHandler(name_path);
            libro.open();
            libro.readFile();
            label1.Text = "Éxito";
        }
    }
}
