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
    public partial class FormTarifas_Calibre : Form
    {
        conexion cn = new conexion();
        public FormTarifas_Calibre()
        {
            InitializeComponent();
        }

        private void FormTarifas_Calibre_Load(object sender, EventArgs e)
        {
            Tarifas_Calibre();
        }
        public void Tarifas_Calibre()
        {
            SqlCommand cmd = new SqlCommand("spTarifaz_Calibre", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView1.DataSource = myds.Tables[0];
            cn.Cerrar();
            dataGridView1.Columns[0].Visible = false;

        }
        public void cargar()
        {
            for (int x = 2; x < dataGridView1.Columns.Count; x++)
            {
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    Update_Tarifas(dataGridView1.Columns[x].Name, dataGridView1.Rows[i].Cells[0].Value.ToString(),Convert.ToSingle(dataGridView1.Rows[i].Cells[x].Value.ToString()));
                 
                }
            }
            MessageBox.Show("Tarifas actualizadas correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        public void Update_Tarifas(string calibre, string tarifa, float monto)
        {
            try
            {
                // linea de comando de sql
                SqlCommand cmd = new SqlCommand("UpdateTarifa_Calibre", cn.getConexion());

                // adhrsion de parametros 
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.Add("@Calibre", SqlDbType.VarChar,10);
                cmd.Parameters.Add("@Tarifa", SqlDbType.SmallInt);
                cmd.Parameters.Add("@monto", SqlDbType.Float);
           


                cmd.Parameters.Add("@msg", SqlDbType.VarChar, 100);

                // ingreso de parametros

                cmd.Parameters["@Calibre"].Value = calibre;
                cmd.Parameters["@Tarifa"].Value = tarifa;
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
        private void btn_ingresar_Click(object sender, EventArgs e)
        {
            cargar();
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.Value.ToString() != "0")
            {
                e.CellStyle.Format = "0.0000";
            }
        }
    } 
}
