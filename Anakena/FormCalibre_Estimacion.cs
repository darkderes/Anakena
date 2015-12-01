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
    public partial class FormCalibre_Estimacion : Form
    {
        public FormCalibre_Estimacion()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
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
                //  formato();
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

        private void FormCalibre_Estimacion_Load(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {        
                
                e.CellStyle.Format = "P02";           
        }
    }
}
