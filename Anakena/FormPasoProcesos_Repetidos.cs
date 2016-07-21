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
    public partial class FormPasoProcesos_Repetidos : Form
    {
        conexion cn = new conexion();
        int cod_variedad = 0;
        public FormPasoProcesos_Repetidos(int variedad)
        {
            InitializeComponent();
            cod_variedad = variedad;
        }

        private void FormPasoProcesos_Repetidos_Load(object sender, EventArgs e)
        {
            //traer_estimacion();
        }
        public void traer_estimacion()
        {
            SqlCommand cmd = new SqlCommand("spProceso_Repetidos", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.VarChar, 25);
            cmd.Parameters["@variedad"].Value = cod_variedad;
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView1.DataSource = myds.Tables[0];
            cn.Cerrar();

        }
        public void cargar()
        {
            for(int i = 0; i< dataGridView1.RowCount; i++)
            {

                Update_Procesos(Convert.ToInt32(dataGridView1.Rows[i].Cells["id_temporal"].Value.ToString()),Convert.ToDouble(dataGridView1.Rows[i].Cells["monto"].Value.ToString()));
            }
        }
        public void Update_Procesos(int temporal,double monto)
        {
            try
            {
                // linea de comando de sql
                SqlCommand cmd = new SqlCommand("Update_Temporal_Envasado", cn.getConexion());

                // adhesion de parametros 
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.Add("@temporal", SqlDbType.Int);
                cmd.Parameters.Add("@monto", SqlDbType.Decimal);
                cmd.Parameters.Add("@msg", SqlDbType.VarChar, 100);

                // ingreso de parametros

                cmd.Parameters["@temporal"].Value = temporal;
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

        private void button1_Click(object sender, EventArgs e)
        {
            cargar();
            this.Close();
        }
    }
}
