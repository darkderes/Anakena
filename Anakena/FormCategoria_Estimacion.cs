using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace Anakena
{
    public partial class FormCategoria_Estimacion : Form
    {
        public FormCategoria_Estimacion()
        {
            InitializeComponent();
        }

        private void Btn_Examinar_Click(object sender, EventArgs e)
        {
            OpenFileDialog newdialog = new OpenFileDialog();
            newdialog.Filter = "Excel files(*.xlsx)|*.xls";
            if (newdialog.ShowDialog() == DialogResult.OK)
            {
                TxtFile.Text = newdialog.FileName;
                importarDatos(TxtFile.Text);
                if (dataGridView1.Rows.Count > 0)
                {
                    btn_ingresar.Visible = true;
                    Btn_Examinar.Visible = false;
                }
            }
        }

        public void importarDatos(string file)
        {

            OleDbCommand command = new OleDbCommand();
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + file + "';Extended Properties=Excel 8.0;";
            command = new OleDbCommand("select * from [Hoja1$]", conn);
            DataSet myds = new DataSet();
            OleDbDataAdapter adap = new OleDbDataAdapter(command);
            adap.Fill(myds);
            dataGridView1.DataSource = myds.Tables[0];

        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            e.CellStyle.Format = "P02";
        }

        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            string nombrecolumna = dataGridView1.Columns[e.ColumnIndex].Name;
            MessageBox.Show(nombrecolumna);
        }

        private void btn_ingresar_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewColumn  col in dataGridView1.Columns)
            {
                string x = col.Name;
            }
        }
    }
}
