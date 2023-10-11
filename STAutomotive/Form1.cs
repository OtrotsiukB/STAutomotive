using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace STAutomotive
{
    public partial class Form1 : Form
    {
        private Excel.Application excelapp;
        private Excel.Window excelWindow;
        public Form1()
        {
            InitializeComponent();
            openFileDialog1.Filter = "Excel(*.xlsx)|*.xlsx|All files(*.*)|*.*";
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        public void openExcel(String wayToFile)
        {
            excelapp = new Excel.Application();
            excelapp.Visible = true;
            Excel.Workbook excelappworkbooks = excelapp.Workbooks.Open(wayToFile,
                         Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                         Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                         Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                         Type.Missing, Type.Missing);
            Excel.Sheets excelsheets = excelappworkbooks.Worksheets;
            Excel.Worksheet
                        //Получаем ссылку на лист 1
                        excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
            Excel.Range
                        //Выбираем ячейку для вывода A1
                        excelcells = excelworksheet.get_Range("A4", Type.Missing);
            String sStr = Convert.ToString(excelcells.Value2);
            excelapp.Quit();
            MessageBox.Show(sStr);

        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            // получаем выбранный файл
            string filename = openFileDialog1.FileName;
            openExcel(filename);
            
        }
    }
}
