using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.Collections;

namespace Anakena
{
    public partial class FormRealidadVSEstimacion : Form
    {
        conexion ex = new conexion();
        conexionERP cn = new conexionERP();
        public FormRealidadVSEstimacion()
        {
            InitializeComponent();
        }

        private void FormRealidadVSEstimacion_Load(object sender, EventArgs e)
        {
            CmbVariedad();
            CmbTipo.SelectedIndex = 0;
            radioButton1.Checked = true;
           
        }
      
      
        public void CmbVariedad()
        {
            try
            {
                ex.Abrir();
                SqlCommand command = new SqlCommand("spTraerVariedad", ex.getConexion());
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
                ex.Cerrar();
            }
        }
        public void EstimacionKG()
        {
            SqlCommand cmd = new SqlCommand("spTraer_Estimacion", ex.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Int);
            cmd.Parameters["@variedad"].Value = cmb_variedad.SelectedValue.ToString();
            ex.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView1.DataSource = myds.Tables[0];
            ex.Cerrar();

        }
        public void EstimacionKG2()
        {
            SqlCommand cmd = new SqlCommand("spTraer_Estimacion", ex.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Int);
            cmd.Parameters["@variedad"].Value = cmb_variedad.SelectedValue.ToString();
            ex.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView3.DataSource = myds.Tables[0];
            ex.Cerrar();

        }

        public void traerReal()
        {
            try
            {
                SqlCommand cmd = new SqlCommand("spTraer_Estimacion2", ex.getConexion());
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@variedad", SqlDbType.VarChar, 25);
                cmd.Parameters["@variedad"].Value =cmb_variedad.SelectedValue.ToString();
                cn.Abrir();
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                DataSet myds = new DataSet();
                adapter.Fill(myds);
                dataGridView2.DataSource = myds.Tables[0];
                cn.Cerrar();
            }
            catch (Exception e)
            {
                MessageBox.Show("Base de datos sin informacion :" + e, "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }
   
        private void BtnFiltro_Click(object sender, EventArgs e)
        {
        
            try
            {
                if ((cmb_variedad.SelectedValue.ToString() != "0") && (radioButton1.Checked == true))
                {
                    Cursor = Cursors.WaitCursor;
                    EstimacionKG();
                    EstimacionKG2();
                    traerReal();
                    dataGridView1.Columns[0].Visible = false;
                    dataGridView1.Columns[1].Visible = false;
                    dataGridView2.Columns[1].Visible = false;
                    dataGridView2.Columns[2].Visible = false;
                   
                    Cursor = Cursors.Default;
                  for (int x = 3; x < dataGridView3.ColumnCount - 1; x++)
                    {
                        for (int i = 0; i < dataGridView3.RowCount - 1; i++)
                        {
                            try
                            { 
                            dataGridView3.Rows[i].Cells[x].Value = 0;
                            dataGridView3.Rows[i].Cells[x].Value = Math.Round((Convert.ToDouble(dataGridView1.Rows[i].Cells[x].Value) - Convert.ToDouble(dataGridView2.Rows[i].Cells[x].Value)),2);
                                if(Convert.ToDouble(dataGridView3.Rows[i].Cells[x].Value.ToString()) < 0)
                                    {
                                  
                                 
                                }
                            }
                            catch { }
                        }

                    }
                      dataGridView3.Columns[0].Visible = false;
                      dataGridView3.Columns[1].Visible = false;
                }
         
            }
            catch (Exception x)
            {
                MessageBox.Show("se produjo un error al tratar de cargar datos  " + x.ToString());
            }  
        }

        private void dataGridView2_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
         if(dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString() == "")
            {
                dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = 0;
            }
           
          
        }

        private void dataGridView3_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            try
            {
                if (Convert.ToDouble(dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()) < 0)
                {
                    dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.SteelBlue;
                    dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.ForeColor = Color.White;
                }
            }
            catch { }
        }
    }
}
