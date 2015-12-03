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
    public partial class FormCategoria_Estimacion : Form
    {
        conexion cn = new conexion();
        public FormCategoria_Estimacion()
        {
            InitializeComponent();
        }

        private void Btn_Examinar_Click(object sender, EventArgs e)
        {
            OpenFileDialog newdialog = new OpenFileDialog();
            newdialog.Filter = "Excel files(*.xlsx)|*.xls";
            if (newdialog.ShowDialog() == DialogResult.OK)
            {
                TxtFile.Text = newdialog.FileName;
                importarDatos(TxtFile.Text);
                if (dataGridView1.Rows.Count > 0)
                {
                    btn_ingresar.Visible = true;
                    Btn_Examinar.Visible = false;
                }
            }
        }

        public void importarDatos(string file)
        {

            OleDbCommand command = new OleDbCommand();
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + file + "';Extended Properties=Excel 8.0;";
            command = new OleDbCommand("select * from [Hoja1$]", conn);
            DataSet myds = new DataSet();
            OleDbDataAdapter adap = new OleDbDataAdapter(command);
            adap.Fill(myds);
            dataGridView1.DataSource = myds.Tables[0];

        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            e.CellStyle.Format = "P02";
        }

        // Trae la llave priamraria de la tabla categoria
        public string traerId_Categoria(string categoria)
        {
            string id = "";
            SqlCommand cmd = new SqlCommand("spTraerIdCategoria", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@categoria", SqlDbType.VarChar, 25);
            cmd.Parameters["@categoria"].Value = categoria;
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            using (SqlDataReader rsd = cmd.ExecuteReader())
            {
                while (rsd.Read())
                {
                    id = rsd[0].ToString();
                }
            }
            cn.Cerrar();
            return id;
        }
        public void ingresarEstimacionCategoria(string codigo, string categoria, float porcentaje)
        {
            try
            {
                // linea de comando de sql
                SqlCommand cmd = new SqlCommand("spAgregaEstimacionCategoria", cn.getConexion());

                // adhrsion de parametros 
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@Cod_Productor", SqlDbType.Char, 10);
                cmd.Parameters.Add("@Cod_Categoria", SqlDbType.Char, 2);
                cmd.Parameters.Add("@Porcentaje", SqlDbType.Float);


                cmd.Parameters.Add("@msg", SqlDbType.VarChar, 100);

                // ingreso de parametros
                cmd.Parameters["@Cod_Productor"].Value = codigo;
                cmd.Parameters["@Cod_Categoria"].Value = categoria;
                cmd.Parameters["@Porcentaje"].Value = porcentaje;


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
            finally
            {


            }
        }
        private void btn_ingresar_Click(object sender, EventArgs e)
        {  
             for (int x = 1; x < dataGridView1.Columns.Count; x++)
            {
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    ingresarEstimacionCategoria(dataGridView1.Columns[x].Name, traerId_Categoria(dataGridView1.Rows[i].Cells[0].Value.ToString()), Convert.ToSingle(dataGridView1.Rows[i].Cells[x].Value.ToString()));
                }

            }
             MessageBox.Show("Estimacion ingresada correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        
        }
    }
}
