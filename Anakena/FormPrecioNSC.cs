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
    public partial class FormPrecioNSC : Form
    {
        conexion cn = new conexion();
        public FormPrecioNSC()
        {
            InitializeComponent();
        }

        private void FormPrecioNSC_Load(object sender, EventArgs e)
        {
            traer_estimacion();
        }
    
       
        public void traer_estimacion()
        {
            SqlCommand cmd = new SqlCommand("spTraer_PrecioNSC", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView1.DataSource = myds.Tables[0];
            cn.Cerrar();

        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.Value.ToString() != "0")
            {
                e.CellStyle.Format = "##.###";
            }
        }

        private void btn_ingresar_Click(object sender, EventArgs e)
        {
            cargar();
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
        public void Update_Precios(string calibre, float monto)
        {
            try
            {
                // linea de comando de sql
                SqlCommand cmd = new SqlCommand("UpdatePrecio_NSC", cn.getConexion());

                // adhrsion de parametros 
                cmd.CommandType = CommandType.StoredProcedure;

               cmd.Parameters.Add("@producto", SqlDbType.VarChar,50);
               cmd.Parameters.Add("@monto", SqlDbType.Float);
               

                cmd.Parameters.Add("@msg", SqlDbType.VarChar, 100);

                // ingreso de parametros

                cmd.Parameters["@producto"].Value = calibre;
                cmd.Parameters["@monto"].Value = monto;
             


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
    }
}
