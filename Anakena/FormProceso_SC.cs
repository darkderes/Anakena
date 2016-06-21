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
    public partial class FormProceso_SC : Form
    {
        conexion cn = new conexion();
        public FormProceso_SC()
        {
            InitializeComponent();
        }

        private void FormProceso_SC_Load(object sender, EventArgs e)
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
            traer_Tarifa();
        }

        public void traer_Tarifa()
        {
            SqlCommand cmd = new SqlCommand("spTarifas_PSC", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.VarChar, 25);
            cmd.Parameters["@variedad"].Value = cmb_variedad.SelectedValue;
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView1.DataSource = myds.Tables[0];
            dataGridView1.Columns[0].Visible = false;
            cn.Cerrar();

        }

        private void btn_ingresar_Click(object sender, EventArgs e)
        {
            cargar();
        }
        public void cargar()
        {
            for (int x = 2; x < dataGridView1.Columns.Count; x++)
            {
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    Update_Tarifas(dataGridView1.Columns[x].Name, dataGridView1.Rows[i].Cells[0].Value.ToString(), Convert.ToSingle(dataGridView1.Rows[i].Cells[x].Value.ToString()));

                }
            }
            MessageBox.Show("Tarifas actualizadas correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void Update_Tarifas(string calibre, string categoria, float monto)
        {
            try
            {
                // linea de comando de sql
                SqlCommand cmd = new SqlCommand("UpdateTarifa_PSC", cn.getConexion());

                // adhrsion de parametros 
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.Add("@Calibre", SqlDbType.VarChar, 10);
                cmd.Parameters.Add("@Categoria", SqlDbType.Char,2);
                cmd.Parameters.Add("@monto", SqlDbType.Float);
                cmd.Parameters.Add("@variedad", SqlDbType.Char,2);



                cmd.Parameters.Add("@msg", SqlDbType.VarChar, 100);

                // ingreso de parametros

                cmd.Parameters["@Calibre"].Value = calibre;
                cmd.Parameters["@Categoria"].Value = categoria;
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

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.Value.ToString() != "0")
            {
                e.CellStyle.Format = "0.0000";
            }
        }
    }
  
}
