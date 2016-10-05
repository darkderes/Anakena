using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Anakena
{
    public partial class FormResumenEstimacionMasReal : Form
    {
        conexion cn = new conexion();
        public FormResumenEstimacionMasReal()
        {
            InitializeComponent();
        }

        private void FormResumenEstimacionMasReal_Load(object sender, EventArgs e)
        {
            CmbVariedad();
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

        private void BtnFiltro_Click(object sender, EventArgs e)
        {
            if (cmb_variedad.SelectedValue.ToString() != "0")
            {
                Estimacion();
            }
            pictureBox1.Visible = true;
            pictureBox2.Visible = false;
            pictureBox3.Visible = false;
            dataGridView4.Visible = true;
            Btn_Excel.Visible = true;
        }
        public void Estimacion()
        {
            SqlCommand cmd = new SqlCommand("spRealmasEstimado_CalibreCategoria", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Int);
            cmd.Parameters["@variedad"].Value = cmb_variedad.SelectedValue.ToString();
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView4.DataSource = myds.Tables[0];
            cn.Cerrar();
            dataGridView4.Columns[0].Visible = false;

        }

        private void dataGridView4_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.Value.ToString() != "0")
            {
                e.CellStyle.Format = "##,##.00";
            }
        }
        public void exporta_a_excel()
        {
          
            ApplicationClass class2 = new ApplicationClass();
            class2.Application.Workbooks.Add(true);
        
            int num = 0;
            pBar1.Minimum = 1;
            pBar1.Maximum = dataGridView4.RowCount * dataGridView4.ColumnCount;
            pBar1.Value = 1;
            pBar1.Step = 1;

            foreach (DataGridViewColumn column in this.dataGridView4.Columns)
            {
               // class2.Cells[1, 1] = cmb_variedad.Text.Trim();
               
                if(column.Name.ToString() != "cod")
                {
                num++;
                class2.Cells[1, num] = column.Name;
                class2.get_Range("A1", "H1").Interior.ColorIndex = 9;
                class2.get_Range("A1", "H1").Font.ColorIndex = 2;
                class2.get_Range("A10", "H10").Interior.ColorIndex = 56;
                class2.get_Range("A10", "H10").Font.ColorIndex = 2;
                class2.get_Range("H1", "H10").Interior.ColorIndex = 56;
                class2.get_Range("H1", "H10").Font.ColorIndex = 2;
                class2.get_Range("A1", "H1").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
              
                }
            }
            int num2 = 0;
            foreach (DataGridViewRow row in (IEnumerable)this.dataGridView4.Rows)
            {
                num2++;
                num = 0;
                foreach (DataGridViewColumn column2 in this.dataGridView4.Columns)
                {
                    if(column2.Name.ToString() != "cod")
                    {
                    num++;
                    class2.Cells[num2 + 1, num] = row.Cells[column2.Name].Value;
                      
                    pBar1.PerformStep();
                    }
                  
                }
            }
            class2.Visible = true;
          
            ((Worksheet)class2.ActiveSheet).Name = cmb_variedad.Text.Trim();    
            ((Worksheet)class2.ActiveSheet).Activate();
            

        }

        private void Btn_Excel_Click(object sender, EventArgs e)
        {
            pBar1.Visible = true;
            exporta_a_excel();
            pBar1.Visible = false;
        }
    }
}
