using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace Anakena
{
    public partial class FormPor_Var : Form
    {
        conexion cn = new conexion();
        public FormPor_Var()
        {
            InitializeComponent();
        }

        private void FormPor_Var_Load(object sender, EventArgs e)
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

        public void Actualizar_Temporal_Lavado()
        {
            try
            {
            // linea de comando de sql
            SqlCommand cmd = new SqlCommand("spAgregarTemporal_Lavado", cn.getConexion());
            // adhrsion de parametros 
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@Cod_Variedad", SqlDbType.Char, 2);
            cmd.Parameters.Add("@msg", SqlDbType.VarChar, 100);
            // ingreso de parametros      
            cmd.Parameters["@Cod_Variedad"].Value = cmb_variedad.SelectedValue.ToString();
            cmd.Parameters["@msg"].Value = 1;
            //abre coneccion a base de datos
            cn.Abrir();
            cmd.ExecuteNonQuery();
            string sqlMensaje = cmd.Parameters["@msg"].Value.ToString();
            //cierre de coneccion a base de datos
            cn.Cerrar();
            }
            catch (Exception ex)
            {
            //mensaje de error SQL
            MessageBox.Show("Error SQL " + ex, "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }  
       }
        public void Actualizar_Temporal_Envasado()
        {
            try
            {
            // linea de comando de sql
            SqlCommand cmd = new SqlCommand("spAgregarTemporal_ENVASADO", cn.getConexion());
            // adhesion de parametros 
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@Cod_Variedad", SqlDbType.Char, 2);
            cmd.Parameters.Add("@msg", SqlDbType.VarChar, 100);
            // ingreso de parametros
            cmd.Parameters["@Cod_Variedad"].Value = cmb_variedad.SelectedValue.ToString();
            cmd.Parameters["@msg"].Value = 1;
            //abre coneccion a base de datos
            cn.Abrir();
            cmd.ExecuteNonQuery();
            string sqlMensaje = cmd.Parameters["@msg"].Value.ToString();
                //cierre de coneccion a base de datos
            cn.Cerrar();
            }
            catch (Exception ex)
            {
                //mensaje de error SQL
                MessageBox.Show("Error SQL " + ex, "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        public void Actualizar_Temporal_Blanquear()
        {
            try
            {
                // linea de comando de sql
                SqlCommand cmd = new SqlCommand("spAgregarTemporal_Blanquear", cn.getConexion());
                // adhesion de parametros 
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@Cod_Variedad", SqlDbType.Char, 2);
                cmd.Parameters.Add("@msg", SqlDbType.VarChar, 100);
                // ingreso de parametros
                cmd.Parameters["@Cod_Variedad"].Value = cmb_variedad.SelectedValue.ToString();
                cmd.Parameters["@msg"].Value = 1;
                //abre coneccion a base de datos
                cn.Abrir();
                cmd.ExecuteNonQuery();
                string sqlMensaje = cmd.Parameters["@msg"].Value.ToString();
                //cierre de coneccion a base de datos
                cn.Cerrar();
            }
            catch (Exception ex)
            {
                //mensaje de error SQL
                MessageBox.Show("Error SQL " + ex, "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void Actualizar_Temporal_Descarte()
        {
            try
            {
                // linea de comando de sql
                SqlCommand cmd = new SqlCommand("spAgregarTemporal_Descarte", cn.getConexion());
                // adhesion de parametros 
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@Cod_Variedad", SqlDbType.Char, 2);
                cmd.Parameters.Add("@msg", SqlDbType.VarChar, 100);
                // ingreso de parametros
                cmd.Parameters["@Cod_Variedad"].Value = cmb_variedad.SelectedValue.ToString();
                cmd.Parameters["@msg"].Value = 1;
                //abre coneccion a base de datos
                cn.Abrir();
                cmd.ExecuteNonQuery();
                string sqlMensaje = cmd.Parameters["@msg"].Value.ToString();
                //cierre de coneccion a base de datos
                cn.Cerrar();
            }
            catch (Exception ex)
            {
                //mensaje de error SQL
                MessageBox.Show("Error SQL " + ex, "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void Actualizar_Temporal_Exportacion()
        {
            try
            {
                // linea de comando de sql
                SqlCommand cmd = new SqlCommand("spAgregarTemporal_Exportacion", cn.getConexion());
                // adhesion de parametros 
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@Cod_Variedad", SqlDbType.Char, 2);
                cmd.Parameters.Add("@msg", SqlDbType.VarChar, 100);
                // ingreso de parametros
                cmd.Parameters["@Cod_Variedad"].Value = cmb_variedad.SelectedValue.ToString();
                cmd.Parameters["@msg"].Value = 1;
                //abre coneccion a base de datos
                cn.Abrir();
                cmd.ExecuteNonQuery();
                string sqlMensaje = cmd.Parameters["@msg"].Value.ToString();
                //cierre de coneccion a base de datos
                cn.Cerrar();
            }
            catch (Exception ex)
            {
                //mensaje de error SQL
                MessageBox.Show("Error SQL " + ex, "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnFiltro_Click(object sender, EventArgs e)
        {

            try
            {
                Actualizar_Temporal_Lavado();
                Actualizar_Temporal_Envasado();
                Actualizar_Temporal_Blanquear();
                Actualizar_Temporal_Descarte();
                Actualizar_Temporal_Exportacion();
                FormPasoProcesos_Repetidos s = new FormPasoProcesos_Repetidos(Convert.ToInt32(cmb_variedad.SelectedValue.ToString()));
                s.traer_estimacion();
                if (s.dataGridView1.RowCount > 0)
                {
                    s.ShowDialog();
                }
                traer_pre_PorVar();
                traer_pre_PorVar2();
                dataGridView1.Visible = true;
                dataGridView1.Columns[0].Visible = false;
                dataGridView2.Visible = true;
                dataGridView2.Columns[0].Visible = false;
                dataGridView2.Rows[1].Visible = false;
                dataGridView2.Rows[0].Visible = false;
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);
                MessageBox.Show(errorMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                dataGridView1.Visible = false;
                dataGridView2.Visible = false;
            }
            finally
            {
                cn.Cerrar();
            }
  
            for (int i = 2; i < dataGridView1.ColumnCount; i++)
            {
                try
                {
                    dataGridView1.Rows[3].Cells[i].Value = (Convert.ToDouble(dataGridView1.Rows[2].Cells[i].Value.ToString()) - Convert.ToDouble(dataGridView1.Rows[8].Cells[i].Value.ToString())) + Convert.ToDouble(dataGridView1.Rows[3].Cells[i].Value.ToString());
                    dataGridView1.Rows[2].Cells[i].Value = Math.Round((Convert.ToDouble(dataGridView1.Rows[0].Cells[i].Value.ToString()) - Convert.ToDouble(dataGridView1.Rows[1].Cells[i].Value.ToString())), 1);
                    if (dataGridView1.Rows[3].Cells[i].Value.ToString() == "")
                    {
                        dataGridView1.Rows[3].Cells[i].Value = 0;
                    }
                    if (dataGridView1.Rows[4].Cells[i].Value.ToString() == "")
                    {
                        dataGridView1.Rows[4].Cells[i].Value = 0;
                    }
                    if (dataGridView1.Rows[5].Cells[i].Value.ToString() == "")
                    {
                        dataGridView1.Rows[5].Cells[i].Value = 0;
                    }
                    if (dataGridView1.Rows[6].Cells[i].Value.ToString() == "")
                    {
                        dataGridView1.Rows[6].Cells[i].Value = 0;
                    }
                    if (dataGridView1.Rows[7].Cells[i].Value.ToString() == "")
                    {
                        dataGridView1.Rows[7].Cells[i].Value = 0;
                    }
                    dataGridView1.Rows[8].Cells[i].Value = Math.Round((Convert.ToDouble(dataGridView1.Rows[2].Cells[i].Value.ToString())) - (Convert.ToDouble(dataGridView1.Rows[3].Cells[i].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[4].Cells[i].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[5].Cells[i].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[6].Cells[i].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[7].Cells[i].Value.ToString())), 1);
                }
                catch { }
            }
            try
            {
            dataGridView1.Rows[8].Cells[1].Value = "MERMA";
            dataGridView2.Rows[8].Cells[1].Value = "MERMA";
            Pre_Porvar();
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);
                MessageBox.Show(errorMessage, "Error",MessageBoxButtons.OK,MessageBoxIcon.Error);
                dataGridView1.Visible = false;
                dataGridView2.Visible = false;
            }
        }

        public void Pre_Porvar()
        {
            for (int i = 2; i < dataGridView2.ColumnCount; i++)
            {
                try
                {
                    dataGridView2.Rows[3].Cells[i].Value = (Convert.ToDouble(dataGridView2.Rows[2].Cells[i].Value.ToString()) - Convert.ToDouble(dataGridView2.Rows[8].Cells[i].Value.ToString())) + Convert.ToDouble(dataGridView2.Rows[3].Cells[i].Value.ToString());
                    dataGridView2.Rows[2].Cells[i].Value = Math.Round((Convert.ToDouble(dataGridView2.Rows[0].Cells[i].Value.ToString()) - Convert.ToDouble(dataGridView2.Rows[1].Cells[i].Value.ToString())), 1);
                    if (dataGridView2.Rows[3].Cells[i].Value.ToString() == "")
                    {
                        dataGridView2.Rows[3].Cells[i].Value = 0;
                    }
                    if (dataGridView2.Rows[4].Cells[i].Value.ToString() == "")
                    {
                        dataGridView2.Rows[4].Cells[i].Value = 0;
                    }
                    if (dataGridView2.Rows[5].Cells[i].Value.ToString() == "")
                    {
                        dataGridView2.Rows[5].Cells[i].Value = 0;
                    }
                    if (dataGridView2.Rows[6].Cells[i].Value.ToString() == "")
                    {
                        dataGridView2.Rows[6].Cells[i].Value = 0;
                    }
                    if (dataGridView2.Rows[7].Cells[i].Value.ToString() == "")
                    {
                        dataGridView2.Rows[7].Cells[i].Value = 0;
                    }
                    dataGridView2.Rows[8].Cells[i].Value = Math.Round((Convert.ToDouble(dataGridView2.Rows[2].Cells[i].Value.ToString())) - (Convert.ToDouble(dataGridView2.Rows[3].Cells[i].Value.ToString()) + Convert.ToDouble(dataGridView2.Rows[4].Cells[i].Value.ToString()) + Convert.ToDouble(dataGridView2.Rows[5].Cells[i].Value.ToString()) + Convert.ToDouble(dataGridView2.Rows[6].Cells[i].Value.ToString()) + Convert.ToDouble(dataGridView2.Rows[7].Cells[i].Value.ToString())), 1);
                }
                catch { }
            try
            {
                dataGridView2.Rows[3].Cells[i].Value = Math.Round((Convert.ToDouble(dataGridView2.Rows[3].Cells[i].Value.ToString()) / Convert.ToDouble(dataGridView2.Rows[2].Cells[i].Value.ToString())) * 100, 1);
            }
            catch { }
            try
            {
                dataGridView2.Rows[4].Cells[i].Value = Math.Round((Convert.ToDouble(dataGridView2.Rows[4].Cells[i].Value.ToString()) / Convert.ToDouble(dataGridView2.Rows[2].Cells[i].Value.ToString())) * 100, 1);
            }
            catch { }
            try
            {
                dataGridView2.Rows[5].Cells[i].Value = Math.Round((Convert.ToDouble(dataGridView2.Rows[5].Cells[i].Value.ToString()) / Convert.ToDouble(dataGridView2.Rows[2].Cells[i].Value.ToString())) * 100, 1);
            }
            catch { }
            try
            {
                dataGridView2.Rows[6].Cells[i].Value = Math.Round((Convert.ToDouble(dataGridView2.Rows[6].Cells[i].Value.ToString()) / Convert.ToDouble(dataGridView2.Rows[2].Cells[i].Value.ToString())) * 100, 1);
            }
            catch { }
            try
            {
                dataGridView2.Rows[7].Cells[i].Value = Math.Round((Convert.ToDouble(dataGridView2.Rows[7].Cells[i].Value.ToString()) / Convert.ToDouble(dataGridView2.Rows[2].Cells[i].Value.ToString())) * 100, 1);
            }
            catch { }
            try
            {
                dataGridView2.Rows[8].Cells[i].Value = Math.Round((Convert.ToDouble(dataGridView2.Rows[8].Cells[i].Value.ToString()) / Convert.ToDouble(dataGridView2.Rows[2].Cells[i].Value.ToString())) * 100, 1);
            }
            catch { }
            try
            {
                dataGridView2.Rows[2].Cells[i].Value = Math.Round((Convert.ToDouble(dataGridView2.Rows[3].Cells[i].Value.ToString()) + Convert.ToDouble(dataGridView2.Rows[4].Cells[i].Value.ToString()) + Convert.ToDouble(dataGridView2.Rows[5].Cells[i].Value.ToString()) + Convert.ToDouble(dataGridView2.Rows[6].Cells[i].Value.ToString()) + Convert.ToDouble(dataGridView2.Rows[7].Cells[i].Value.ToString()) + Convert.ToDouble(dataGridView2.Rows[8].Cells[i].Value.ToString())), 0);
            }
            catch { }
            }
        }
       
        public void traer_pre_PorVar()
        {
            SqlCommand cmd = new SqlCommand("sp_PrePorVar", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.VarChar, 25);
            cmd.Parameters["@variedad"].Value = cmb_variedad.SelectedValue.ToString();
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView1.DataSource = myds.Tables[0];
            cn.Cerrar();
        }

        public void traer_pre_PorVar2()
        {
            SqlCommand cmd = new SqlCommand("sp_PrePorVar", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.VarChar, 25);
            cmd.Parameters["@variedad"].Value = cmb_variedad.SelectedValue.ToString();
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView2.DataSource = myds.Tables[0];
            dataGridView2.Rows[0].Visible = false;
            dataGridView2.Rows[1].Visible = false;
            cn.Cerrar();
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString() == "")
            {
                dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = 0;
            }
        }
        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            MessageBox.Show(e.RowIndex.ToString() + "/" + e.ColumnIndex.ToString());
        }
        private void dataGridView2_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString() == "")
            {
                dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = 0;
            }
        }
        private void dataGridView1_Scroll(object sender, ScrollEventArgs e)
        {
            dataGridView1.Columns[1].Frozen = true;
            dataGridView2.HorizontalScrollingOffset = e.NewValue;
        }

        private void dataGridView2_Scroll(object sender, ScrollEventArgs e)
        {
            dataGridView2.Columns[1].Frozen = true;
            dataGridView1.HorizontalScrollingOffset = e.NewValue;
        }
    }
  
}
