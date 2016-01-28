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
    public partial class FormBusquedaProductor : Form
    {
        conexion cn = new conexion();
        public string nombre;
        public string codigo;
        public FormBusquedaProductor()
        {
            InitializeComponent();
        }

        private void FormBusquedaProductor_Load(object sender, EventArgs e)
        {

        }
        private void filtrar(string filtro)
        {
            try
            {
                SqlCommand cmd = new SqlCommand("spFiltroProductores", cn.getConexion());
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@filtro", SqlDbType.VarChar,10);
                cmd.Parameters["@filtro"].Value = filtro;
                cn.Abrir();
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                DataSet myds = new DataSet();
                myds = new DataSet();
                adapter.Fill(myds);
                dataGridView1.DataSource = myds.Tables[0];
                cn.Cerrar();
            }
            catch (Exception ex)
            {
                //mensaje de error SQL
                MessageBox.Show("Error SQL " + ex, "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {


            }

        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == Convert.ToChar(Keys.Enter))
            { 
            filtrar(textBox1.Text);
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            nombre = dataGridView1.Rows[e.RowIndex].Cells["Nom_Productor"].Value.ToString();
            codigo = dataGridView1.Rows[e.RowIndex].Cells["Cod_Productor"].Value.ToString();
            this.Close();
        }
    }
}
