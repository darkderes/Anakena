using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Anakena
{
    public partial class FormPor_NSC : Form
    {
        conexion cn = new conexion();
        public FormPor_NSC()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Traer_Por_NSC(200,390,427,464,501);
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Planilla_Pos", 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Excel._Worksheet xlWorksheet_36_EXTRA = (Excel._Worksheet)xlWorkbook.Sheets[1];
            Excel._Worksheet xlWorksheet_36_CATI = (Excel._Worksheet)xlWorkbook.Sheets[2];
            Excel._Worksheet xlWorksheet_36_CATII = (Excel._Worksheet)xlWorkbook.Sheets[3];
            Excel._Worksheet xlWorksheet_36_CATIII = (Excel._Worksheet)xlWorkbook.Sheets[4];
            Excel._Worksheet xlWorksheet_36_B_CATIII = (Excel._Worksheet)xlWorkbook.Sheets[5];
            Excel._Worksheet xlWorksheet_34_36_EXTRA = (Excel._Worksheet)xlWorkbook.Sheets[6];
            Excel._Worksheet xlWorksheet_34_36_CATI = (Excel._Worksheet)xlWorkbook.Sheets[7];
            Excel._Worksheet xlWorksheet_34_36_CATII = (Excel._Worksheet)xlWorkbook.Sheets[8];
            Excel._Worksheet xlWorksheet_34_36_CATIII = (Excel._Worksheet)xlWorkbook.Sheets[9];
            Excel._Worksheet xlWorksheet_34_36_B_CATIII = (Excel._Worksheet)xlWorkbook.Sheets[10];
            Excel._Worksheet xlWorksheet_32_34_EXTRA = (Excel._Worksheet)xlWorkbook.Sheets[11];
            Excel._Worksheet xlWorksheet_32_34_CATI = (Excel._Worksheet)xlWorkbook.Sheets[12];
            Excel._Worksheet xlWorksheet_32_34_CATII = (Excel._Worksheet)xlWorkbook.Sheets[13];
            Excel._Worksheet xlWorksheet_32_34_CATIII = (Excel._Worksheet)xlWorkbook.Sheets[14];
            Excel._Worksheet xlWorksheet_32_34_B_CATIII = (Excel._Worksheet)xlWorkbook.Sheets[15];
            Excel._Worksheet xlWorksheet_30_32_EXTRA = (Excel._Worksheet)xlWorkbook.Sheets[16];
            Excel._Worksheet xlWorksheet_30_32_CATI = (Excel._Worksheet)xlWorkbook.Sheets[17];
            Excel._Worksheet xlWorksheet_30_32_CATII = (Excel._Worksheet)xlWorkbook.Sheets[18];
            Excel._Worksheet xlWorksheet_30_32_CATIII = (Excel._Worksheet)xlWorkbook.Sheets[19];
            Excel._Worksheet xlWorksheet_30_32_B_CATIII = (Excel._Worksheet)xlWorkbook.Sheets[20];
            Excel._Worksheet xlWorksheet_28_30_EXTRA = (Excel._Worksheet)xlWorkbook.Sheets[21];
            Excel._Worksheet xlWorksheet_28_30_CATI = (Excel._Worksheet)xlWorkbook.Sheets[22];
            Excel._Worksheet xlWorksheet_28_30_CATII = (Excel._Worksheet)xlWorkbook.Sheets[23];
            Excel._Worksheet xlWorksheet_28_30_CATIII = (Excel._Worksheet)xlWorkbook.Sheets[24];
            Excel._Worksheet xlWorksheet_28_30_B_CATIII = (Excel._Worksheet)xlWorkbook.Sheets[25];
            // Excel.Range xlRange = xlWorksheet.UsedRange;
            xlWorksheet_36_EXTRA.Cells[1, 1] = cmb_variedad.Text;
            int y = 2;
            for (int i = 0;i<dataGridView1.RowCount;i++)
            {
                Double ACUM = 0;
                DateTime date = DateTime.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString());
                xlWorksheet_36_EXTRA.Cells[3, y] = date.ToString("dd/MM");
                xlWorksheet_36_EXTRA.Cells[4, y] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                xlWorksheet_36_EXTRA.Cells[5, y] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                xlWorksheet_36_EXTRA.Cells[6, y] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                xlWorksheet_36_EXTRA.Cells[7, y] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                xlWorksheet_36_EXTRA.Cells[8, y] = dataGridView1.Rows[i].Cells[5].Value.ToString();
                xlWorksheet_36_EXTRA.Cells[11, y] = dataGridView1.Rows[i].Cells[7].Value.ToString();
                xlWorksheet_36_EXTRA.Cells[12, y] = dataGridView1.Rows[i].Cells[8].Value.ToString();
                xlWorksheet_36_EXTRA.Cells[13, y] = dataGridView1.Rows[i].Cells[9].Value.ToString();
                xlWorksheet_36_EXTRA.Cells[14, y] = dataGridView1.Rows[i].Cells[10].Value.ToString();
                xlWorksheet_36_EXTRA.Cells[15, y] = dataGridView1.Rows[i].Cells[11].Value.ToString();
                xlWorksheet_36_EXTRA.Cells[16, y] = dataGridView1.Rows[i].Cells[12].Value.ToString();
                ACUM = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString())+ Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value.ToString())+ Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value.ToString())+ Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString())+ Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value.ToString())+ Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value.ToString());


                xlWorksheet_36_EXTRA.Cells[19, y] = ACUM;
                xlWorksheet_36_EXTRA.Cells[20, y] = Math.Round( (ACUM / Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString()))* 100,1);
                xlWorksheet_36_EXTRA.Cells[21, y] = Math.Round((Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) / ACUM) * 100,1);
                y++;
            }
            xlWorksheet_36_CATI.Cells[1, 1] = cmb_variedad.Text;
            y = 2;
            Traer_Por_NSC(201,391,428,465,502);
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                Double ACUM = 0;
                DateTime date = DateTime.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString());
                xlWorksheet_36_CATI.Cells[3, y] = date.ToString("dd/MM");
                xlWorksheet_36_CATI.Cells[4, y] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                xlWorksheet_36_CATI.Cells[5, y] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                xlWorksheet_36_CATI.Cells[6, y] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                xlWorksheet_36_CATI.Cells[7, y] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                xlWorksheet_36_CATI.Cells[8, y] = dataGridView1.Rows[i].Cells[5].Value.ToString();
                xlWorksheet_36_CATI.Cells[11, y] = dataGridView1.Rows[i].Cells[7].Value.ToString();
                xlWorksheet_36_CATI.Cells[12, y] = dataGridView1.Rows[i].Cells[8].Value.ToString();
                xlWorksheet_36_CATI.Cells[13, y] = dataGridView1.Rows[i].Cells[9].Value.ToString();
                xlWorksheet_36_CATI.Cells[14, y] = dataGridView1.Rows[i].Cells[10].Value.ToString();
                xlWorksheet_36_CATI.Cells[15, y] = dataGridView1.Rows[i].Cells[11].Value.ToString();
                xlWorksheet_36_CATI.Cells[16, y] = dataGridView1.Rows[i].Cells[12].Value.ToString();
                ACUM = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value.ToString());
                xlWorksheet_36_CATI.Cells.Cells[19, y] = ACUM;
                xlWorksheet_36_CATI.Cells.Cells[20, y] = Math.Round((ACUM / Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString())) * 100, 1);
                xlWorksheet_36_CATI.Cells.Cells[21, y] = Math.Round((Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) / ACUM) * 100, 1);
                y++;
            }        
            xlWorksheet_36_CATII.Cells[1, 1] = cmb_variedad.Text;
            y = 2;
            Traer_Por_NSC(202, 392, 429, 466, 503);
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                DateTime date = DateTime.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString());
                Double ACUM = 0;
                xlWorksheet_36_CATII.Cells[3, y] = date.ToString("dd/MM");
                xlWorksheet_36_CATII.Cells[4, y] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                xlWorksheet_36_CATII.Cells[5, y] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                xlWorksheet_36_CATII.Cells[6, y] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                xlWorksheet_36_CATII.Cells[7, y] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                xlWorksheet_36_CATII.Cells[8, y] = dataGridView1.Rows[i].Cells[5].Value.ToString();
                xlWorksheet_36_CATII.Cells[11, y] = dataGridView1.Rows[i].Cells[7].Value.ToString();
                xlWorksheet_36_CATII.Cells[12, y] = dataGridView1.Rows[i].Cells[8].Value.ToString();
                xlWorksheet_36_CATII.Cells[13, y] = dataGridView1.Rows[i].Cells[9].Value.ToString();
                xlWorksheet_36_CATII.Cells[14, y] = dataGridView1.Rows[i].Cells[10].Value.ToString();
                xlWorksheet_36_CATII.Cells[15, y] = dataGridView1.Rows[i].Cells[11].Value.ToString();
                xlWorksheet_36_CATII.Cells[16, y] = dataGridView1.Rows[i].Cells[12].Value.ToString();
                ACUM = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value.ToString());
                xlWorksheet_36_CATII.Cells.Cells[19, y] = ACUM;
                xlWorksheet_36_CATII.Cells.Cells[20, y] = Math.Round((ACUM / Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString())) * 100, 1);
                xlWorksheet_36_CATII.Cells.Cells[21, y] = Math.Round((Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) / ACUM) * 100, 1);
                y++;
            }
            xlWorksheet_36_CATIII.Cells[1, 1] = cmb_variedad.Text;
            y = 2;
            Traer_Por_NSC(203, 393, 430, 467, 504);
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                DateTime date = DateTime.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString());
                Double ACUM = 0;
                xlWorksheet_36_CATIII.Cells[3, y] = date.ToString("dd/MM");
                xlWorksheet_36_CATIII.Cells[4, y] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                xlWorksheet_36_CATIII.Cells[5, y] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                xlWorksheet_36_CATIII.Cells[6, y] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                xlWorksheet_36_CATIII.Cells[7, y] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                xlWorksheet_36_CATIII.Cells[8, y] = dataGridView1.Rows[i].Cells[5].Value.ToString();
                xlWorksheet_36_CATIII.Cells[11, y] = dataGridView1.Rows[i].Cells[7].Value.ToString();
                xlWorksheet_36_CATIII.Cells[12, y] = dataGridView1.Rows[i].Cells[8].Value.ToString();
                xlWorksheet_36_CATIII.Cells[13, y] = dataGridView1.Rows[i].Cells[9].Value.ToString();
                xlWorksheet_36_CATIII.Cells[14, y] = dataGridView1.Rows[i].Cells[10].Value.ToString();
                xlWorksheet_36_CATIII.Cells[15, y] = dataGridView1.Rows[i].Cells[11].Value.ToString();
                xlWorksheet_36_CATIII.Cells[16, y] = dataGridView1.Rows[i].Cells[12].Value.ToString();
                ACUM = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value.ToString());
                xlWorksheet_36_CATIII.Cells.Cells[19, y] = ACUM;
                xlWorksheet_36_CATIII.Cells.Cells[20, y] = Math.Round((ACUM / Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString())) * 100, 1);
                xlWorksheet_36_CATIII.Cells.Cells[21, y] = Math.Round((Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) / ACUM) * 100, 1);
                y++;
            }
            xlWorksheet_36_B_CATIII.Cells[1, 1] = cmb_variedad.Text;
            y = 2;
            Traer_Por_NSC(204, 394, 431, 468, 505);
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                DateTime date = DateTime.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString());
                Double ACUM = 0;
                xlWorksheet_36_B_CATIII.Cells[3, y] = date.ToString("dd/MM");
                xlWorksheet_36_B_CATIII.Cells[4, y] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                xlWorksheet_36_B_CATIII.Cells[5, y] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                xlWorksheet_36_B_CATIII.Cells[6, y] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                xlWorksheet_36_B_CATIII.Cells[7, y] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                xlWorksheet_36_B_CATIII.Cells[8, y] = dataGridView1.Rows[i].Cells[5].Value.ToString();
                xlWorksheet_36_B_CATIII.Cells[11, y] = dataGridView1.Rows[i].Cells[7].Value.ToString();
                xlWorksheet_36_B_CATIII.Cells[12, y] = dataGridView1.Rows[i].Cells[8].Value.ToString();
                xlWorksheet_36_B_CATIII.Cells[13, y] = dataGridView1.Rows[i].Cells[9].Value.ToString();
                xlWorksheet_36_B_CATIII.Cells[14, y] = dataGridView1.Rows[i].Cells[10].Value.ToString();
                xlWorksheet_36_B_CATIII.Cells[15, y] = dataGridView1.Rows[i].Cells[11].Value.ToString();
                xlWorksheet_36_B_CATIII.Cells[16, y] = dataGridView1.Rows[i].Cells[12].Value.ToString();
                ACUM = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value.ToString());
                xlWorksheet_36_B_CATIII.Cells.Cells[19, y] = ACUM;
                xlWorksheet_36_B_CATIII.Cells.Cells[20, y] = Math.Round((ACUM / Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString())) * 100, 1);
                xlWorksheet_36_B_CATIII.Cells.Cells[21, y] = Math.Round((Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) / ACUM) * 100, 1);
                y++;
            }
            xlWorksheet_34_36_EXTRA.Cells[1, 1] = cmb_variedad.Text;
            y = 2;
            Traer_Por_NSC(205, 395, 432, 469, 506);
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                DateTime date = DateTime.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString());
                Double ACUM = 0;
                xlWorksheet_34_36_EXTRA.Cells[3, y] = date.ToString("dd/MM");
                xlWorksheet_34_36_EXTRA.Cells[4, y] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                xlWorksheet_34_36_EXTRA.Cells[5, y] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                xlWorksheet_34_36_EXTRA.Cells[6, y] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                xlWorksheet_34_36_EXTRA.Cells[7, y] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                xlWorksheet_34_36_EXTRA.Cells[8, y] = dataGridView1.Rows[i].Cells[5].Value.ToString();
                xlWorksheet_34_36_EXTRA.Cells[11, y] = dataGridView1.Rows[i].Cells[7].Value.ToString();
                xlWorksheet_34_36_EXTRA.Cells[12, y] = dataGridView1.Rows[i].Cells[8].Value.ToString();
                xlWorksheet_34_36_EXTRA.Cells[13, y] = dataGridView1.Rows[i].Cells[9].Value.ToString();
                xlWorksheet_34_36_EXTRA.Cells[14, y] = dataGridView1.Rows[i].Cells[10].Value.ToString();
                xlWorksheet_34_36_EXTRA.Cells[15, y] = dataGridView1.Rows[i].Cells[11].Value.ToString();
                xlWorksheet_34_36_EXTRA.Cells[16, y] = dataGridView1.Rows[i].Cells[12].Value.ToString();
                ACUM = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value.ToString());
                xlWorksheet_34_36_EXTRA.Cells.Cells[19, y] = ACUM;
                xlWorksheet_34_36_EXTRA.Cells.Cells[20, y] = Math.Round((ACUM / Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString())) * 100, 1);
                xlWorksheet_34_36_EXTRA.Cells.Cells[21, y] = Math.Round((Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) / ACUM) * 100, 1);
                y++;
            }
            xlWorksheet_34_36_CATI.Cells[1, 1] = cmb_variedad.Text;
            y = 2;
            Traer_Por_NSC(206, 396, 433, 470, 507);
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                DateTime date = DateTime.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString());
                Double ACUM = 0;
                xlWorksheet_34_36_CATI.Cells[3, y] = date.ToString("dd/MM");
                xlWorksheet_34_36_CATI.Cells[4, y] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                xlWorksheet_34_36_CATI.Cells[5, y] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                xlWorksheet_34_36_CATI.Cells[6, y] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                xlWorksheet_34_36_CATI.Cells[7, y] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                xlWorksheet_34_36_CATI.Cells[8, y] = dataGridView1.Rows[i].Cells[5].Value.ToString();
                xlWorksheet_34_36_CATI.Cells[11, y] = dataGridView1.Rows[i].Cells[7].Value.ToString();
                xlWorksheet_34_36_CATI.Cells[12, y] = dataGridView1.Rows[i].Cells[8].Value.ToString();
                xlWorksheet_34_36_CATI.Cells[13, y] = dataGridView1.Rows[i].Cells[9].Value.ToString();
                xlWorksheet_34_36_CATI.Cells[14, y] = dataGridView1.Rows[i].Cells[10].Value.ToString();
                xlWorksheet_34_36_CATI.Cells[15, y] = dataGridView1.Rows[i].Cells[11].Value.ToString();
                xlWorksheet_34_36_CATI.Cells[16, y] = dataGridView1.Rows[i].Cells[12].Value.ToString();
                ACUM = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value.ToString());
                xlWorksheet_34_36_CATI.Cells.Cells[19, y] = ACUM;
                xlWorksheet_34_36_CATI.Cells.Cells[20, y] = Math.Round((ACUM / Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString())) * 100, 1);
                xlWorksheet_34_36_CATI.Cells.Cells[21, y] = Math.Round((Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) / ACUM) * 100, 1);
                y++;
            }
            xlWorksheet_34_36_CATII.Cells[1, 1] = cmb_variedad.Text;
            y = 2;
            Traer_Por_NSC(207, 397, 434, 471, 508);
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                DateTime date = DateTime.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString());
                Double ACUM = 0;
                xlWorksheet_34_36_CATII.Cells[3, y] = date.ToString("dd/MM");
                xlWorksheet_34_36_CATII.Cells[4, y] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                xlWorksheet_34_36_CATII.Cells[5, y] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                xlWorksheet_34_36_CATII.Cells[6, y] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                xlWorksheet_34_36_CATII.Cells[7, y] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                xlWorksheet_34_36_CATII.Cells[8, y] = dataGridView1.Rows[i].Cells[5].Value.ToString();
                xlWorksheet_34_36_CATII.Cells[11, y] = dataGridView1.Rows[i].Cells[7].Value.ToString();
                xlWorksheet_34_36_CATII.Cells[12, y] = dataGridView1.Rows[i].Cells[8].Value.ToString();
                xlWorksheet_34_36_CATII.Cells[13, y] = dataGridView1.Rows[i].Cells[9].Value.ToString();
                xlWorksheet_34_36_CATII.Cells[14, y] = dataGridView1.Rows[i].Cells[10].Value.ToString();
                xlWorksheet_34_36_CATII.Cells[15, y] = dataGridView1.Rows[i].Cells[11].Value.ToString();
                xlWorksheet_34_36_CATII.Cells[16, y] = dataGridView1.Rows[i].Cells[12].Value.ToString();
                ACUM = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value.ToString());
                xlWorksheet_34_36_CATII.Cells.Cells[19, y] = ACUM;
                xlWorksheet_34_36_CATII.Cells.Cells[20, y] = Math.Round((ACUM / Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString())) * 100, 1);
                xlWorksheet_34_36_CATII.Cells.Cells[21, y] = Math.Round((Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) / ACUM) * 100, 1);
                y++;
            }
            xlWorksheet_34_36_CATIII.Cells[1, 1] = cmb_variedad.Text;
            y = 2;
            Traer_Por_NSC(208, 398, 435, 472, 509);
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                DateTime date = DateTime.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString());
                Double ACUM = 0;
                xlWorksheet_34_36_CATIII.Cells[3, y] = date.ToString("dd/MM");
                xlWorksheet_34_36_CATIII.Cells[4, y] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                xlWorksheet_34_36_CATIII.Cells[5, y] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                xlWorksheet_34_36_CATIII.Cells[6, y] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                xlWorksheet_34_36_CATIII.Cells[7, y] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                xlWorksheet_34_36_CATIII.Cells[8, y] = dataGridView1.Rows[i].Cells[5].Value.ToString();
                xlWorksheet_34_36_CATIII.Cells[11, y] = dataGridView1.Rows[i].Cells[7].Value.ToString();
                xlWorksheet_34_36_CATIII.Cells[12, y] = dataGridView1.Rows[i].Cells[8].Value.ToString();
                xlWorksheet_34_36_CATIII.Cells[13, y] = dataGridView1.Rows[i].Cells[9].Value.ToString();
                xlWorksheet_34_36_CATIII.Cells[14, y] = dataGridView1.Rows[i].Cells[10].Value.ToString();
                xlWorksheet_34_36_CATIII.Cells[15, y] = dataGridView1.Rows[i].Cells[11].Value.ToString();
                xlWorksheet_34_36_CATIII.Cells[16, y] = dataGridView1.Rows[i].Cells[12].Value.ToString();
                ACUM = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value.ToString());
                xlWorksheet_34_36_CATIII.Cells.Cells[19, y] = ACUM;
                xlWorksheet_34_36_CATIII.Cells.Cells[20, y] = Math.Round((ACUM / Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString())) * 100, 1);
                xlWorksheet_34_36_CATIII.Cells.Cells[21, y] = Math.Round((Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) / ACUM) * 100, 1);
                y++;
            }
            xlWorksheet_34_36_B_CATIII.Cells[1, 1] = cmb_variedad.Text;
            y = 2;
            Traer_Por_NSC(209, 399, 436, 473, 510);
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                DateTime date = DateTime.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString());
                Double ACUM = 0;
                xlWorksheet_34_36_B_CATIII.Cells[3, y] = date.ToString("dd/MM");
                xlWorksheet_34_36_B_CATIII.Cells[4, y] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                xlWorksheet_34_36_B_CATIII.Cells[5, y] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                xlWorksheet_34_36_B_CATIII.Cells[6, y] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                xlWorksheet_34_36_B_CATIII.Cells[7, y] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                xlWorksheet_34_36_B_CATIII.Cells[8, y] = dataGridView1.Rows[i].Cells[5].Value.ToString();
                xlWorksheet_34_36_B_CATIII.Cells[11, y] = dataGridView1.Rows[i].Cells[7].Value.ToString();
                xlWorksheet_34_36_B_CATIII.Cells[12, y] = dataGridView1.Rows[i].Cells[8].Value.ToString();
                xlWorksheet_34_36_B_CATIII.Cells[13, y] = dataGridView1.Rows[i].Cells[9].Value.ToString();
                xlWorksheet_34_36_B_CATIII.Cells[14, y] = dataGridView1.Rows[i].Cells[10].Value.ToString();
                xlWorksheet_34_36_B_CATIII.Cells[15, y] = dataGridView1.Rows[i].Cells[11].Value.ToString();
                xlWorksheet_34_36_B_CATIII.Cells[16, y] = dataGridView1.Rows[i].Cells[12].Value.ToString();
                ACUM = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value.ToString());
                xlWorksheet_34_36_B_CATIII.Cells.Cells[19, y] = ACUM;
                xlWorksheet_34_36_B_CATIII.Cells.Cells[20, y] = Math.Round((ACUM / Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString())) * 100, 1);
                xlWorksheet_34_36_B_CATIII.Cells.Cells[21, y] = Math.Round((Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) / ACUM) * 100, 1);
                y++;
            }
            xlWorksheet_32_34_EXTRA.Cells[1, 1] = cmb_variedad.Text;
            y = 2;
            Traer_Por_NSC(210,400, 437, 474, 511);
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                DateTime date = DateTime.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString());
                Double ACUM = 0;
                xlWorksheet_32_34_EXTRA.Cells[3, y] = date.ToString("dd/MM");
                xlWorksheet_32_34_EXTRA.Cells[4, y] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                xlWorksheet_32_34_EXTRA.Cells[5, y] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                xlWorksheet_32_34_EXTRA.Cells[6, y] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                xlWorksheet_32_34_EXTRA.Cells[7, y] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                xlWorksheet_32_34_EXTRA.Cells[8, y] = dataGridView1.Rows[i].Cells[5].Value.ToString();
                xlWorksheet_32_34_EXTRA.Cells[11, y] = dataGridView1.Rows[i].Cells[7].Value.ToString();
                xlWorksheet_32_34_EXTRA.Cells[12, y] = dataGridView1.Rows[i].Cells[8].Value.ToString();
                xlWorksheet_32_34_EXTRA.Cells[13, y] = dataGridView1.Rows[i].Cells[9].Value.ToString();
                xlWorksheet_32_34_EXTRA.Cells[14, y] = dataGridView1.Rows[i].Cells[10].Value.ToString();
                xlWorksheet_32_34_EXTRA.Cells[15, y] = dataGridView1.Rows[i].Cells[11].Value.ToString();
                xlWorksheet_32_34_EXTRA.Cells[16, y] = dataGridView1.Rows[i].Cells[12].Value.ToString();
                ACUM = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value.ToString());
                xlWorksheet_32_34_EXTRA.Cells.Cells[19, y] = ACUM;
                xlWorksheet_32_34_EXTRA.Cells.Cells[20, y] = Math.Round((ACUM / Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString())) * 100, 1);
                xlWorksheet_32_34_EXTRA.Cells.Cells[21, y] = Math.Round((Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) / ACUM) * 100, 1);
                y++;
            }
            xlWorksheet_32_34_CATI.Cells[1, 1] = cmb_variedad.Text;
            y = 2;
            Traer_Por_NSC(211, 401, 438, 475, 512);
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                DateTime date = DateTime.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString());
                Double ACUM = 0;
                xlWorksheet_32_34_CATI.Cells[3, y] = date.ToString("dd/MM");
                xlWorksheet_32_34_CATI.Cells[4, y] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                xlWorksheet_32_34_CATI.Cells[5, y] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                xlWorksheet_32_34_CATI.Cells[6, y] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                xlWorksheet_32_34_CATI.Cells[7, y] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                xlWorksheet_32_34_CATI.Cells[8, y] = dataGridView1.Rows[i].Cells[5].Value.ToString();
                xlWorksheet_32_34_CATI.Cells[11, y] = dataGridView1.Rows[i].Cells[7].Value.ToString();
                xlWorksheet_32_34_CATI.Cells[12, y] = dataGridView1.Rows[i].Cells[8].Value.ToString();
                xlWorksheet_32_34_CATI.Cells[13, y] = dataGridView1.Rows[i].Cells[9].Value.ToString();
                xlWorksheet_32_34_CATI.Cells[14, y] = dataGridView1.Rows[i].Cells[10].Value.ToString();
                xlWorksheet_32_34_CATI.Cells[15, y] = dataGridView1.Rows[i].Cells[11].Value.ToString();
                xlWorksheet_32_34_CATI.Cells[16, y] = dataGridView1.Rows[i].Cells[12].Value.ToString();
                ACUM = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value.ToString());
                xlWorksheet_32_34_CATI.Cells.Cells[19, y] = ACUM;
                xlWorksheet_32_34_CATI.Cells.Cells[20, y] = Math.Round((ACUM / Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString())) * 100, 1);
                xlWorksheet_32_34_CATI.Cells.Cells[21, y] = Math.Round((Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) / ACUM) * 100, 1);
                y++;
            }
            xlWorksheet_32_34_CATII.Cells[1, 1] = cmb_variedad.Text;
            y = 2;
            Traer_Por_NSC(212, 402, 439, 476, 513);
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                DateTime date = DateTime.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString());
                Double ACUM = 0;
                xlWorksheet_32_34_CATII.Cells[3, y] = date.ToString("dd/MM");
                xlWorksheet_32_34_CATII.Cells[4, y] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                xlWorksheet_32_34_CATII.Cells[5, y] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                xlWorksheet_32_34_CATII.Cells[6, y] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                xlWorksheet_32_34_CATII.Cells[7, y] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                xlWorksheet_32_34_CATII.Cells[8, y] = dataGridView1.Rows[i].Cells[5].Value.ToString();
                xlWorksheet_32_34_CATII.Cells[11, y] = dataGridView1.Rows[i].Cells[7].Value.ToString();
                xlWorksheet_32_34_CATII.Cells[12, y] = dataGridView1.Rows[i].Cells[8].Value.ToString();
                xlWorksheet_32_34_CATII.Cells[13, y] = dataGridView1.Rows[i].Cells[9].Value.ToString();
                xlWorksheet_32_34_CATII.Cells[14, y] = dataGridView1.Rows[i].Cells[10].Value.ToString();
                xlWorksheet_32_34_CATII.Cells[15, y] = dataGridView1.Rows[i].Cells[11].Value.ToString();
                xlWorksheet_32_34_CATII.Cells[16, y] = dataGridView1.Rows[i].Cells[12].Value.ToString();
                ACUM = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value.ToString());
                xlWorksheet_32_34_CATII.Cells.Cells[19, y] = ACUM;
                xlWorksheet_32_34_CATII.Cells.Cells[20, y] = Math.Round((ACUM / Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString())) * 100, 1);
                xlWorksheet_32_34_CATII.Cells.Cells[21, y] = Math.Round((Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) / ACUM) * 100, 1);
                y++;
            }
            xlWorksheet_32_34_CATIII.Cells[1, 1] = cmb_variedad.Text;
            y = 2;
            Traer_Por_NSC(213, 403, 440, 477, 514);
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                DateTime date = DateTime.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString());
                Double ACUM = 0;
                xlWorksheet_32_34_CATIII.Cells[3, y] = date.ToString("dd/MM");
                xlWorksheet_32_34_CATIII.Cells[4, y] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                xlWorksheet_32_34_CATIII.Cells[5, y] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                xlWorksheet_32_34_CATIII.Cells[6, y] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                xlWorksheet_32_34_CATIII.Cells[7, y] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                xlWorksheet_32_34_CATIII.Cells[8, y] = dataGridView1.Rows[i].Cells[5].Value.ToString();
                xlWorksheet_32_34_CATIII.Cells[11, y] = dataGridView1.Rows[i].Cells[7].Value.ToString();
                xlWorksheet_32_34_CATIII.Cells[12, y] = dataGridView1.Rows[i].Cells[8].Value.ToString();
                xlWorksheet_32_34_CATIII.Cells[13, y] = dataGridView1.Rows[i].Cells[9].Value.ToString();
                xlWorksheet_32_34_CATIII.Cells[14, y] = dataGridView1.Rows[i].Cells[10].Value.ToString();
                xlWorksheet_32_34_CATIII.Cells[15, y] = dataGridView1.Rows[i].Cells[11].Value.ToString();
                xlWorksheet_32_34_CATIII.Cells[16, y] = dataGridView1.Rows[i].Cells[12].Value.ToString();
                ACUM = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value.ToString());
                xlWorksheet_32_34_CATIII.Cells.Cells[19, y] = ACUM;
                xlWorksheet_32_34_CATIII.Cells.Cells[20, y] = Math.Round((ACUM / Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString())) * 100, 1);
                xlWorksheet_32_34_CATIII.Cells.Cells[21, y] = Math.Round((Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) / ACUM) * 100, 1);
                y++;
            }
            xlWorksheet_32_34_B_CATIII.Cells[1, 1] = cmb_variedad.Text;
            y = 2;
            Traer_Por_NSC(214, 404, 441, 478, 515);
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                DateTime date = DateTime.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString());
                Double ACUM = 0;
                xlWorksheet_32_34_B_CATIII.Cells[3, y] = date.ToString("dd/MM");
                xlWorksheet_32_34_B_CATIII.Cells[4, y] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                xlWorksheet_32_34_B_CATIII.Cells[5, y] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                xlWorksheet_32_34_B_CATIII.Cells[6, y] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                xlWorksheet_32_34_B_CATIII.Cells[7, y] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                xlWorksheet_32_34_B_CATIII.Cells[8, y] = dataGridView1.Rows[i].Cells[5].Value.ToString();
                xlWorksheet_32_34_B_CATIII.Cells[11, y] = dataGridView1.Rows[i].Cells[7].Value.ToString();
                xlWorksheet_32_34_B_CATIII.Cells[12, y] = dataGridView1.Rows[i].Cells[8].Value.ToString();
                xlWorksheet_32_34_B_CATIII.Cells[13, y] = dataGridView1.Rows[i].Cells[9].Value.ToString();
                xlWorksheet_32_34_B_CATIII.Cells[14, y] = dataGridView1.Rows[i].Cells[10].Value.ToString();
                xlWorksheet_32_34_B_CATIII.Cells[15, y] = dataGridView1.Rows[i].Cells[11].Value.ToString();
                xlWorksheet_32_34_B_CATIII.Cells[16, y] = dataGridView1.Rows[i].Cells[12].Value.ToString();
                ACUM = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value.ToString());
                xlWorksheet_32_34_B_CATIII.Cells.Cells[19, y] = ACUM;
                xlWorksheet_32_34_B_CATIII.Cells.Cells[20, y] = Math.Round((ACUM / Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString())) * 100, 1);
                xlWorksheet_32_34_B_CATIII.Cells.Cells[21, y] = Math.Round((Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) / ACUM) * 100, 1);
                y++;
            }
            xlWorksheet_30_32_EXTRA.Cells[1, 1] = cmb_variedad.Text;
            y = 2;
            Traer_Por_NSC(215, 405, 442, 479, 516);
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                DateTime date = DateTime.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString());
                Double ACUM = 0;
                xlWorksheet_30_32_EXTRA.Cells[3, y] = date.ToString("dd/MM");
                xlWorksheet_30_32_EXTRA.Cells[4, y] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                xlWorksheet_30_32_EXTRA.Cells[5, y] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                xlWorksheet_30_32_EXTRA.Cells[6, y] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                xlWorksheet_30_32_EXTRA.Cells[7, y] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                xlWorksheet_30_32_EXTRA.Cells[8, y] = dataGridView1.Rows[i].Cells[5].Value.ToString();
                xlWorksheet_30_32_EXTRA.Cells[11, y] = dataGridView1.Rows[i].Cells[7].Value.ToString();
                xlWorksheet_30_32_EXTRA.Cells[12, y] = dataGridView1.Rows[i].Cells[8].Value.ToString();
                xlWorksheet_30_32_EXTRA.Cells[13, y] = dataGridView1.Rows[i].Cells[9].Value.ToString();
                xlWorksheet_30_32_EXTRA.Cells[14, y] = dataGridView1.Rows[i].Cells[10].Value.ToString();
                xlWorksheet_30_32_EXTRA.Cells[15, y] = dataGridView1.Rows[i].Cells[11].Value.ToString();
                xlWorksheet_30_32_EXTRA.Cells[16, y] = dataGridView1.Rows[i].Cells[12].Value.ToString();
                ACUM = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value.ToString());
                xlWorksheet_30_32_EXTRA.Cells.Cells[19, y] = ACUM;
                xlWorksheet_30_32_EXTRA.Cells.Cells[20, y] = Math.Round((ACUM / Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString())) * 100, 1);
                xlWorksheet_30_32_EXTRA.Cells.Cells[21, y] = Math.Round((Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) / ACUM) * 100, 1);
                y++;
            }
            xlWorksheet_30_32_CATI.Cells[1, 1] = cmb_variedad.Text;
            y = 2;
            Traer_Por_NSC(216, 406, 443, 480, 517);
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                DateTime date = DateTime.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString());
                Double ACUM = 0;
                xlWorksheet_30_32_CATI.Cells[3, y] = date.ToString("dd/MM");
                xlWorksheet_30_32_CATI.Cells[4, y] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                xlWorksheet_30_32_CATI.Cells[5, y] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                xlWorksheet_30_32_CATI.Cells[6, y] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                xlWorksheet_30_32_CATI.Cells[7, y] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                xlWorksheet_30_32_CATI.Cells[8, y] = dataGridView1.Rows[i].Cells[5].Value.ToString();
                xlWorksheet_30_32_CATI.Cells[11, y] = dataGridView1.Rows[i].Cells[7].Value.ToString();
                xlWorksheet_30_32_CATI.Cells[12, y] = dataGridView1.Rows[i].Cells[8].Value.ToString();
                xlWorksheet_30_32_CATI.Cells[13, y] = dataGridView1.Rows[i].Cells[9].Value.ToString();
                xlWorksheet_30_32_CATI.Cells[14, y] = dataGridView1.Rows[i].Cells[10].Value.ToString();
                xlWorksheet_30_32_CATI.Cells[15, y] = dataGridView1.Rows[i].Cells[11].Value.ToString();
                xlWorksheet_30_32_CATI.Cells[16, y] = dataGridView1.Rows[i].Cells[12].Value.ToString();
                ACUM = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value.ToString());
                xlWorksheet_30_32_CATI.Cells.Cells[19, y] = ACUM;
                xlWorksheet_30_32_CATI.Cells.Cells[20, y] = Math.Round((ACUM / Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString())) * 100, 1);
                xlWorksheet_30_32_CATI.Cells.Cells[21, y] = Math.Round((Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) / ACUM) * 100, 1);
                y++;
            }
            xlWorksheet_30_32_CATII.Cells[1, 1] = cmb_variedad.Text;
            y = 2;
            Traer_Por_NSC(217, 407, 444, 481, 518);
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                DateTime date = DateTime.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString());
                Double ACUM = 0;
                xlWorksheet_30_32_CATII.Cells[3, y] = date.ToString("dd/MM");
                xlWorksheet_30_32_CATII.Cells[4, y] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                xlWorksheet_30_32_CATII.Cells[5, y] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                xlWorksheet_30_32_CATII.Cells[6, y] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                xlWorksheet_30_32_CATII.Cells[7, y] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                xlWorksheet_30_32_CATII.Cells[8, y] = dataGridView1.Rows[i].Cells[5].Value.ToString();
                xlWorksheet_30_32_CATII.Cells[11, y] = dataGridView1.Rows[i].Cells[7].Value.ToString();
                xlWorksheet_30_32_CATII.Cells[12, y] = dataGridView1.Rows[i].Cells[8].Value.ToString();
                xlWorksheet_30_32_CATII.Cells[13, y] = dataGridView1.Rows[i].Cells[9].Value.ToString();
                xlWorksheet_30_32_CATII.Cells[14, y] = dataGridView1.Rows[i].Cells[10].Value.ToString();
                xlWorksheet_30_32_CATII.Cells[15, y] = dataGridView1.Rows[i].Cells[11].Value.ToString();
                xlWorksheet_30_32_CATII.Cells[16, y] = dataGridView1.Rows[i].Cells[12].Value.ToString();
                ACUM = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value.ToString());
                xlWorksheet_30_32_CATII.Cells.Cells[19, y] = ACUM;
                xlWorksheet_30_32_CATII.Cells.Cells[20, y] = Math.Round((ACUM / Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString())) * 100, 1);
                xlWorksheet_30_32_CATII.Cells.Cells[21, y] = Math.Round((Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) / ACUM) * 100, 1);
                y++;
            }
            xlWorksheet_30_32_CATIII.Cells[1, 1] = cmb_variedad.Text;
            y = 2;
            Traer_Por_NSC(218, 408, 445, 482, 519);
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                DateTime date = DateTime.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString());
                Double ACUM = 0;
                xlWorksheet_30_32_CATIII.Cells[3, y] = date.ToString("dd/MM");
                xlWorksheet_30_32_CATIII.Cells[4, y] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                xlWorksheet_30_32_CATIII.Cells[5, y] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                xlWorksheet_30_32_CATIII.Cells[6, y] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                xlWorksheet_30_32_CATIII.Cells[7, y] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                xlWorksheet_30_32_CATIII.Cells[8, y] = dataGridView1.Rows[i].Cells[5].Value.ToString();
                xlWorksheet_30_32_CATIII.Cells[11, y] = dataGridView1.Rows[i].Cells[7].Value.ToString();
                xlWorksheet_30_32_CATIII.Cells[12, y] = dataGridView1.Rows[i].Cells[8].Value.ToString();
                xlWorksheet_30_32_CATIII.Cells[13, y] = dataGridView1.Rows[i].Cells[9].Value.ToString();
                xlWorksheet_30_32_CATIII.Cells[14, y] = dataGridView1.Rows[i].Cells[10].Value.ToString();
                xlWorksheet_30_32_CATIII.Cells[15, y] = dataGridView1.Rows[i].Cells[11].Value.ToString();
                xlWorksheet_30_32_CATIII.Cells[16, y] = dataGridView1.Rows[i].Cells[12].Value.ToString();
                ACUM = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value.ToString());
                xlWorksheet_30_32_CATIII.Cells.Cells[19, y] = ACUM;
                xlWorksheet_30_32_CATIII.Cells.Cells[20, y] = Math.Round((ACUM / Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString())) * 100, 1);
                xlWorksheet_30_32_CATIII.Cells.Cells[21, y] = Math.Round((Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) / ACUM) * 100, 1);
                y++;
            }
            xlWorksheet_30_32_B_CATIII.Cells[1, 1] = cmb_variedad.Text;
            y = 2;
            Traer_Por_NSC(219, 409, 446, 483, 519);
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                DateTime date = DateTime.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString());
                Double ACUM = 0;
                xlWorksheet_30_32_B_CATIII.Cells[3, y] = date.ToString("dd/MM");
                xlWorksheet_30_32_B_CATIII.Cells[4, y] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                xlWorksheet_30_32_B_CATIII.Cells[5, y] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                xlWorksheet_30_32_B_CATIII.Cells[6, y] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                xlWorksheet_30_32_B_CATIII.Cells[7, y] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                xlWorksheet_30_32_B_CATIII.Cells[8, y] = dataGridView1.Rows[i].Cells[5].Value.ToString();
                xlWorksheet_30_32_B_CATIII.Cells[11, y] = dataGridView1.Rows[i].Cells[7].Value.ToString();
                xlWorksheet_30_32_B_CATIII.Cells[12, y] = dataGridView1.Rows[i].Cells[8].Value.ToString();
                xlWorksheet_30_32_B_CATIII.Cells[13, y] = dataGridView1.Rows[i].Cells[9].Value.ToString();
                xlWorksheet_30_32_B_CATIII.Cells[14, y] = dataGridView1.Rows[i].Cells[10].Value.ToString();
                xlWorksheet_30_32_B_CATIII.Cells[15, y] = dataGridView1.Rows[i].Cells[11].Value.ToString();
                xlWorksheet_30_32_B_CATIII.Cells[16, y] = dataGridView1.Rows[i].Cells[12].Value.ToString();
                ACUM = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value.ToString());
                xlWorksheet_30_32_B_CATIII.Cells.Cells[19, y] = ACUM;
                xlWorksheet_30_32_B_CATIII.Cells.Cells[20, y] = Math.Round((ACUM / Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString())) * 100, 1);
                xlWorksheet_30_32_B_CATIII.Cells.Cells[21, y] = Math.Round((Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) / ACUM) * 100, 1);
                y++;
            }
            xlWorksheet_28_30_EXTRA.Cells[1, 1] = cmb_variedad.Text;
            y = 2;
            Traer_Por_NSC(220, 410, 447, 484, 520);
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                DateTime date = DateTime.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString());
                Double ACUM = 0;
                xlWorksheet_28_30_EXTRA.Cells[3, y] = date.ToString("dd/MM");
                xlWorksheet_28_30_EXTRA.Cells[4, y] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                xlWorksheet_28_30_EXTRA.Cells[5, y] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                xlWorksheet_28_30_EXTRA.Cells[6, y] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                xlWorksheet_28_30_EXTRA.Cells[7, y] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                xlWorksheet_28_30_EXTRA.Cells[8, y] = dataGridView1.Rows[i].Cells[5].Value.ToString();
                xlWorksheet_28_30_EXTRA.Cells[11, y] = dataGridView1.Rows[i].Cells[7].Value.ToString();
                xlWorksheet_28_30_EXTRA.Cells[12, y] = dataGridView1.Rows[i].Cells[8].Value.ToString();
                xlWorksheet_28_30_EXTRA.Cells[13, y] = dataGridView1.Rows[i].Cells[9].Value.ToString();
                xlWorksheet_28_30_EXTRA.Cells[14, y] = dataGridView1.Rows[i].Cells[10].Value.ToString();
                xlWorksheet_28_30_EXTRA.Cells[15, y] = dataGridView1.Rows[i].Cells[11].Value.ToString();
                xlWorksheet_28_30_EXTRA.Cells[16, y] = dataGridView1.Rows[i].Cells[12].Value.ToString();
                ACUM = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value.ToString());
                xlWorksheet_28_30_EXTRA.Cells.Cells[19, y] = ACUM;
                xlWorksheet_28_30_EXTRA.Cells.Cells[20, y] = Math.Round((ACUM / Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString())) * 100, 1);
                xlWorksheet_28_30_EXTRA.Cells.Cells[21, y] = Math.Round((Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) / ACUM) * 100, 1);
                y++;
            }
            xlWorksheet_28_30_CATI.Cells[1, 1] = cmb_variedad.Text;
            y = 2;
            Traer_Por_NSC(221, 411, 448, 485, 521);
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                DateTime date = DateTime.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString());
                Double ACUM = 0;
                xlWorksheet_28_30_CATI.Cells[3, y] = date.ToString("dd/MM");
                xlWorksheet_28_30_CATI.Cells[4, y] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                xlWorksheet_28_30_CATI.Cells[5, y] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                xlWorksheet_28_30_CATI.Cells[6, y] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                xlWorksheet_28_30_CATI.Cells[7, y] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                xlWorksheet_28_30_CATI.Cells[8, y] = dataGridView1.Rows[i].Cells[5].Value.ToString();
                xlWorksheet_28_30_CATI.Cells[11, y] = dataGridView1.Rows[i].Cells[7].Value.ToString();
                xlWorksheet_28_30_CATI.Cells[12, y] = dataGridView1.Rows[i].Cells[8].Value.ToString();
                xlWorksheet_28_30_CATI.Cells[13, y] = dataGridView1.Rows[i].Cells[9].Value.ToString();
                xlWorksheet_28_30_CATI.Cells[14, y] = dataGridView1.Rows[i].Cells[10].Value.ToString();
                xlWorksheet_28_30_CATI.Cells[15, y] = dataGridView1.Rows[i].Cells[11].Value.ToString();
                xlWorksheet_28_30_CATI.Cells[16, y] = dataGridView1.Rows[i].Cells[12].Value.ToString();
                ACUM = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value.ToString());
                xlWorksheet_28_30_CATI.Cells.Cells[19, y] = ACUM;
                xlWorksheet_28_30_CATI.Cells.Cells[20, y] = Math.Round((ACUM / Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString())) * 100, 1);
                xlWorksheet_28_30_CATI.Cells.Cells[21, y] = Math.Round((Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) / ACUM) * 100, 1);
                y++;
            }
            xlWorksheet_28_30_CATII.Cells[1, 1] = cmb_variedad.Text;
            y = 2;
            Traer_Por_NSC(222, 412, 449, 486, 522);
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                DateTime date = DateTime.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString());
                Double ACUM = 0;
                xlWorksheet_28_30_CATII.Cells[3, y] = date.ToString("dd/MM");
                xlWorksheet_28_30_CATII.Cells[4, y] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                xlWorksheet_28_30_CATII.Cells[5, y] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                xlWorksheet_28_30_CATII.Cells[6, y] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                xlWorksheet_28_30_CATII.Cells[7, y] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                xlWorksheet_28_30_CATII.Cells[8, y] = dataGridView1.Rows[i].Cells[5].Value.ToString();
                xlWorksheet_28_30_CATII.Cells[11, y] = dataGridView1.Rows[i].Cells[7].Value.ToString();
                xlWorksheet_28_30_CATII.Cells[12, y] = dataGridView1.Rows[i].Cells[8].Value.ToString();
                xlWorksheet_28_30_CATII.Cells[13, y] = dataGridView1.Rows[i].Cells[9].Value.ToString();
                xlWorksheet_28_30_CATII.Cells[14, y] = dataGridView1.Rows[i].Cells[10].Value.ToString();
                xlWorksheet_28_30_CATII.Cells[15, y] = dataGridView1.Rows[i].Cells[11].Value.ToString();
                xlWorksheet_28_30_CATII.Cells[16, y] = dataGridView1.Rows[i].Cells[12].Value.ToString();
                ACUM = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value.ToString());
                xlWorksheet_28_30_CATII.Cells.Cells[19, y] = ACUM;
                xlWorksheet_28_30_CATII.Cells.Cells[20, y] = Math.Round((ACUM / Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString())) * 100, 1);
                xlWorksheet_28_30_CATII.Cells.Cells[21, y] = Math.Round((Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) / ACUM) * 100, 1);
                y++;
            }
            xlWorksheet_28_30_CATIII.Cells[1, 1] = cmb_variedad.Text;
            y = 2;
            Traer_Por_NSC(223, 413, 450, 487, 523);
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                DateTime date = DateTime.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString());
                Double ACUM = 0;
                xlWorksheet_28_30_CATIII.Cells[3, y] = date.ToString("dd/MM");
                xlWorksheet_28_30_CATIII.Cells[4, y] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                xlWorksheet_28_30_CATIII.Cells[5, y] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                xlWorksheet_28_30_CATIII.Cells[6, y] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                xlWorksheet_28_30_CATIII.Cells[7, y] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                xlWorksheet_28_30_CATIII.Cells[8, y] = dataGridView1.Rows[i].Cells[5].Value.ToString();
                xlWorksheet_28_30_CATIII.Cells[11, y] = dataGridView1.Rows[i].Cells[7].Value.ToString();
                xlWorksheet_28_30_CATIII.Cells[12, y] = dataGridView1.Rows[i].Cells[8].Value.ToString();
                xlWorksheet_28_30_CATIII.Cells[13, y] = dataGridView1.Rows[i].Cells[9].Value.ToString();
                xlWorksheet_28_30_CATIII.Cells[14, y] = dataGridView1.Rows[i].Cells[10].Value.ToString();
                xlWorksheet_28_30_CATIII.Cells[15, y] = dataGridView1.Rows[i].Cells[11].Value.ToString();
                xlWorksheet_28_30_CATIII.Cells[16, y] = dataGridView1.Rows[i].Cells[12].Value.ToString();
                ACUM = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value.ToString());
                xlWorksheet_28_30_CATIII.Cells.Cells[19, y] = ACUM;
                xlWorksheet_28_30_CATIII.Cells.Cells[20, y] = Math.Round((ACUM / Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString())) * 100, 1);
                xlWorksheet_28_30_CATIII.Cells.Cells[21, y] = Math.Round((Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) / ACUM) * 100, 1);
                y++;
            }
            xlWorksheet_28_30_B_CATIII.Cells[1, 1] = cmb_variedad.Text;
            y = 2;
            Traer_Por_NSC(224, 404, 441, 478, 515);
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                DateTime date = DateTime.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString());
                Double ACUM = 0;
                xlWorksheet_28_30_B_CATIII.Cells[3, y] = date.ToString("dd/MM");
                xlWorksheet_28_30_B_CATIII.Cells[4, y] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                xlWorksheet_28_30_B_CATIII.Cells[5, y] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                xlWorksheet_28_30_B_CATIII.Cells[6, y] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                xlWorksheet_28_30_B_CATIII.Cells[7, y] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                xlWorksheet_28_30_B_CATIII.Cells[8, y] = dataGridView1.Rows[i].Cells[5].Value.ToString();
                xlWorksheet_28_30_B_CATIII.Cells[11, y] = dataGridView1.Rows[i].Cells[7].Value.ToString();
                xlWorksheet_28_30_B_CATIII.Cells[12, y] = dataGridView1.Rows[i].Cells[8].Value.ToString();
                xlWorksheet_28_30_B_CATIII.Cells[13, y] = dataGridView1.Rows[i].Cells[9].Value.ToString();
                xlWorksheet_28_30_B_CATIII.Cells[14, y] = dataGridView1.Rows[i].Cells[10].Value.ToString();
                xlWorksheet_28_30_B_CATIII.Cells[15, y] = dataGridView1.Rows[i].Cells[11].Value.ToString();
                xlWorksheet_28_30_B_CATIII.Cells[16, y] = dataGridView1.Rows[i].Cells[12].Value.ToString();
                ACUM = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value.ToString());
                xlWorksheet_28_30_B_CATIII.Cells.Cells[19, y] = ACUM;
                xlWorksheet_28_30_B_CATIII.Cells.Cells[20, y] = Math.Round((ACUM / Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString())) * 100, 1);
                xlWorksheet_28_30_B_CATIII.Cells.Cells[21, y] = Math.Round((Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString()) / ACUM) * 100, 1);
                y++;
            }

            xlApp.Visible = true;
            this.Close();
        }
        public void Traer_Por_NSC(int producto,int producto_halves,int Large_Pieces,int Medium_Pieces,int Small_Pieces)
        {
            SqlCommand cmd = new SqlCommand("spTraerPor_NSC", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@Producto", SqlDbType.Int);
            cmd.Parameters["@Producto"].Value = producto;
            cmd.Parameters.Add("@Cod_Variedad", SqlDbType.Int);
            cmd.Parameters["@Cod_Variedad"].Value = cmb_variedad.SelectedValue.ToString();
            cmd.Parameters.Add("@Cod_Halves", SqlDbType.Int);
            cmd.Parameters["@Cod_Halves"].Value = producto_halves;
            cmd.Parameters.Add("@Cod_LargePieces", SqlDbType.Int);
            cmd.Parameters["@Cod_LargePieces"].Value = Large_Pieces;
            cmd.Parameters.Add("@Cod_MediumPieces", SqlDbType.Int);
            cmd.Parameters["@Cod_MediumPieces"].Value = Medium_Pieces;
            cmd.Parameters.Add("@Cod_SmallPieces", SqlDbType.Int);
            cmd.Parameters["@Cod_SmallPieces"].Value = Small_Pieces;
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView1.DataSource = myds.Tables[0];
            cn.Cerrar();
        }
        public void CmbVariedad()
        {
            try
            {
                cn.Abrir();
                SqlCommand command = new SqlCommand("spTraerVariedad", cn.getConexion());
                SqlDataAdapter myAdapter = new SqlDataAdapter(command);
                command.CommandType = CommandType.StoredProcedure;
                DataSet myds = new DataSet();
                myAdapter.Fill(myds, "Variedad");
                cmb_variedad.Refresh();
                cmb_variedad.DataSource = myds.Tables["Variedad"].DefaultView;
                cmb_variedad.DisplayMember = "Des_Variedad";
                cmb_variedad.ValueMember = "Cod_Variedad";
            }
            catch
            {
                MessageBox.Show("Ocurrio un problema al cargar variedades", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cn.Cerrar();
            }
        }
        private void FormPor_NSC_Load(object sender, EventArgs e)
        {
            CmbVariedad();
        }
    }
}
