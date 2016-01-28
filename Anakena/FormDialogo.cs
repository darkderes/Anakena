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
    public partial class FormDialogo : Form
    {
        conexion cn = new conexion();
        public FormDialogo()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FormDialogo_Load(object sender, EventArgs e)
        {
            
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

        private void btn_guardar_Click(object sender, EventArgs e)
        {
            string x = "";
            if (Convert.ToDouble(TxtEstimacion.Value.ToString())+Convert.ToDouble(TxtReal.Value.ToString()) == 100)
            {           
                if (Convert.ToDouble(txtProductor.Value.ToString()) < 10)
                {
                    x = "0" + txtProductor.Value.ToString();                
                }
                else
                {
                    x = txtProductor.Value.ToString();
                }
                try
                {
                    if(cmb_variedad.SelectedIndex != 0)
                    { 
                     ingresarFactor(x, cmb_variedad.SelectedValue.ToString(), Convert.ToDouble(TxtEstimacion.Value.ToString()), Convert.ToDouble(TxtReal.Value.ToString()));
                      MessageBox.Show("Valores ingresados perfectamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);
                      this.Close();
                    }
                    else
                    {
                        MessageBox.Show("debe seleccioar variedad", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                 
                }
                
                catch { }
            }
            else
            {
                MessageBox.Show("la suma de los valores real y estimacion deben ser igual a 100", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }
        public void ingresarFactor(string codigo, string variedad, double porcentaje_est, double porcentaje_real)
        {
            try
            {
                // linea de comando de sql
                SqlCommand cmd = new SqlCommand("spAgregaFactores", cn.getConexion());

                // adhrsion de parametros 
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@Cod_Productor", SqlDbType.Char, 10);
                cmd.Parameters.Add("@Cod_Variedad", SqlDbType.Char, 2);
                cmd.Parameters.Add("@Porcentaje_estimacion", SqlDbType.Float);
                cmd.Parameters.Add("@Porcentaje_real", SqlDbType.Float);


                cmd.Parameters.Add("@msg", SqlDbType.VarChar, 100);

                // ingreso de parametros
                cmd.Parameters["@Cod_Productor"].Value = codigo;

                cmd.Parameters["@Cod_Variedad"].Value = variedad;
                cmd.Parameters["@Porcentaje_estimacion"].Value = porcentaje_est;
                cmd.Parameters["@Porcentaje_real"].Value = porcentaje_real;


                cmd.Parameters["@msg"].Value = 1;
                //abre coneccion a base de datos
                cn.Abrir();
                cmd.ExecuteNonQuery();
                string sqlMensaje = cmd.Parameters["@msg"].Value.ToString();

                //cierre de coneccion a base de datos

            }
            catch (Exception ex)
            {
                //mensaje de error SQL
                MessageBox.Show("Error SQL " + ex, "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cn.Cerrar();
            }
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void TxtEstimacion_KeyPress(object sender, KeyPressEventArgs e)
        {
            if(e.KeyChar == Convert.ToChar(Keys.Enter))
            {
                TxtReal.Value = 100 - Convert.ToDecimal(TxtEstimacion.Value.ToString());
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            FormBusquedaProductor s = new FormBusquedaProductor();
            s.ShowDialog();
            txtProductor.Value = Convert.ToDecimal(s.codigo);
        }

        private void TxtReal_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == Convert.ToChar(Keys.Enter))
            {
                TxtEstimacion.Value = 100 - Convert.ToDecimal(TxtReal.Value.ToString());
            }
        }
    }
}
