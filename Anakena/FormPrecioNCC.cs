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

namespace Anakena
{
    public partial class FormPrecioNCC : Form
    {
        conexion cn = new conexion();
        public FormPrecioNCC()
        {
            InitializeComponent();
        }

        private void FormPrecioNCC_Load(object sender, EventArgs e)
        {
            CmbVariedad();
        }
        public void traer_estimacion()
        {
            SqlCommand cmd = new SqlCommand("spTraer_PrecioNCC", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.VarChar, 25);
            cmd.Parameters["@variedad"].Value = cmb_variedad.SelectedValue;
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

        private void BtnFiltro_Click(object sender, EventArgs e)
        {
            traer_estimacion();
            if (dataGridView1.RowCount>0)
            {
                dataGridView1.Visible = true;
                pictureBox1.Visible = true;
                pictureBox2.Visible = false;
                pictureBox3.Visible = false;
                btn_ingresar.Visible = true;
            }
            else
            {
                MessageBox.Show("No se encontraron datos ingresados para esta variedad", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        public void Update_Precios(string calibre, float monto)
        {
            try
            {
                // linea de comando de sql
                SqlCommand cmd = new SqlCommand("UpdatePrecio_NCC", cn.getConexion());

                // adhrsion de parametros 
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.Add("@Calibre", SqlDbType.VarChar, 10);
                cmd.Parameters.Add("@monto", SqlDbType.Float);
                cmd.Parameters.Add("@variedad", SqlDbType.Char, 2);

                cmd.Parameters.Add("@msg", SqlDbType.VarChar, 100);

                // ingreso de parametros

                cmd.Parameters["@Calibre"].Value = calibre;
                cmd.Parameters["@monto"].Value = monto;
                cmd.Parameters["@variedad"].Value = cmb_variedad.SelectedValue.ToString();


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
        public void cargar()
        {
            try
            {
             
                    for (int i = 0; i < dataGridView1.RowCount; i++)
                    {
                      
                        Update_Precios(dataGridView1.Rows[i].Cells[0].Value.ToString(), Convert.ToSingle(dataGridView1.Rows[i].Cells[1].Value.ToString()));

                    }
                
                MessageBox.Show("Precios actualizados correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);
                traer_estimacion();
            }
            catch { }
        }
        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.Value.ToString() != "0")
            {
                e.CellStyle.Format = "##.##";
            }
        }

        private void btn_ingresar_Click(object sender, EventArgs e)
        {
            cargar();
        }
    }
}
