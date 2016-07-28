using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Anakena
{
    public partial class FormPor_NCC : Form
    {
        conexion cn = new conexion();
        public FormPor_NCC()
        {
            InitializeComponent();
        }

        private void FormPor_NCC_Load(object sender, EventArgs e)
        {

        }
        public void importarDatosSERR(string file)
        {
            try
            {
                OleDbCommand command = new OleDbCommand();
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + file + "';Extended Properties=Excel 8.0;";
                command = new OleDbCommand("select * from [SERR$]", conn);
                DataSet myds = new DataSet();
                OleDbDataAdapter adap = new OleDbDataAdapter(command);
                adap.Fill(myds);
                dataGridView1.DataSource = myds.Tables[0];
            }
            catch
            {
                MessageBox.Show("Error al tratar cargar Hoja 'SERR' verifique que existe en el archivo", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

        }
        public void importarDatosCHANDLER(string file)
        {
            try
            {
                OleDbCommand command = new OleDbCommand();
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + file + "';Extended Properties=Excel 8.0;";
                command = new OleDbCommand("select * from [CHANDLER$]", conn);
                DataSet myds = new DataSet();
                OleDbDataAdapter adap = new OleDbDataAdapter(command);
                adap.Fill(myds);
                dataGridView2.DataSource = myds.Tables[0];
            }
            catch
            {
                MessageBox.Show("Error al tratar cargar Hoja 'CHANDLER' verifique que existe en el archivo", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        public void importarDatosHOWARD(string file)
        {
            try
            {
                OleDbCommand command = new OleDbCommand();
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + file + "';Extended Properties=Excel 8.0;";
                command = new OleDbCommand("select * from [HOWARD$]", conn);
                DataSet myds = new DataSet();
                OleDbDataAdapter adap = new OleDbDataAdapter(command);
                adap.Fill(myds);
                dataGridView3.DataSource = myds.Tables[0];
            }
            catch
            {
                MessageBox.Show("Error al tratar cargar Hoja 'HOWARDS' verifique que existe en el archivo", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void importarDatosTULARE(string file)
        {
            try
            {
                OleDbCommand command = new OleDbCommand();
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + file + "';Extended Properties=Excel 8.0;";
                command = new OleDbCommand("select * from [TULARE$]", conn);
                DataSet myds = new DataSet();
                OleDbDataAdapter adap = new OleDbDataAdapter(command);
                adap.Fill(myds);
                dataGridView4.DataSource = myds.Tables[0];
            }
            catch
            {
                MessageBox.Show("Error al tratar cargar Hoja 'TULARES' verifique que existe en el archivo", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        public void importarDatosSUNDLAND(string file)
        {
            try
            {
                OleDbCommand command = new OleDbCommand();
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + file + "';Extended Properties=Excel 8.0;";
                command = new OleDbCommand("select * from [SUNDLAND$]", conn);
                DataSet myds = new DataSet();
                OleDbDataAdapter adap = new OleDbDataAdapter(command);
                adap.Fill(myds);
                dataGridView5.DataSource = myds.Tables[0];
            }
            catch
            {
                MessageBox.Show("Error al tratar cargar Hoja 'SUNDLAND' verifique que existe en el archivo", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void importarDatosHARTLEY(string file)
        {
            try
            {
                OleDbCommand command = new OleDbCommand();
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + file + "';Extended Properties=Excel 8.0;";
                command = new OleDbCommand("select * from [HARTLEY$]", conn);
                DataSet myds = new DataSet();
                OleDbDataAdapter adap = new OleDbDataAdapter(command);
                adap.Fill(myds);
                dataGridView6.DataSource = myds.Tables[0];
            }
            catch
            {
                MessageBox.Show("Error al tratar cargar Hoja 'HARTLEY' verifique que existe en el archivo", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void importarDatosSEMILLA(string file)
        {
            try
            {
                OleDbCommand command = new OleDbCommand();
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + file + "';Extended Properties=Excel 8.0;";
                command = new OleDbCommand("select * from [SEMILLA$]", conn);
                DataSet myds = new DataSet();
                OleDbDataAdapter adap = new OleDbDataAdapter(command);
                adap.Fill(myds);
                dataGridView7.DataSource = myds.Tables[0];
            }
            catch
            {
                MessageBox.Show("Error al tratar cargar Hoja 'SEMILLA' verifique que existe en el archivo", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void importarDatosVINA(string file)
        {
            try
            {
                OleDbCommand command = new OleDbCommand();
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + file + "';Extended Properties=Excel 8.0;";
                command = new OleDbCommand("select * from [VINA$]", conn);
                DataSet myds = new DataSet();
                OleDbDataAdapter adap = new OleDbDataAdapter(command);
                adap.Fill(myds);
                dataGridView8.DataSource = myds.Tables[0];
            }
            catch
            {
                MessageBox.Show("Error al tratar cargar Hoja 'VINA' verifique que existe en el archivo", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void importarDatosVINA_MEJ(string file)
        {
            try
            {
                OleDbCommand command = new OleDbCommand();
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + file + "';Extended Properties=Excel 8.0;";
                command = new OleDbCommand("select * from [VINA MEJORADA$]", conn);
                DataSet myds = new DataSet();
                OleDbDataAdapter adap = new OleDbDataAdapter(command);
                adap.Fill(myds);
                dataGridView9.DataSource = myds.Tables[0];
            }
            catch
            {
                MessageBox.Show("Error al tratar cargar Hoja 'VINA MEJORADA' verifique que existe en el archivo", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void importarDatosFRANQUETTE(string file)
        {
            try
            {
                OleDbCommand command = new OleDbCommand();
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + file + "';Extended Properties=Excel 8.0;";
                command = new OleDbCommand("select * from [FRANQUETTE$]", conn);
                DataSet myds = new DataSet();
                OleDbDataAdapter adap = new OleDbDataAdapter(command);
                adap.Fill(myds);
                dataGridView10.DataSource = myds.Tables[0];
            }
            catch
            {
                MessageBox.Show("Error al tratar cargar Hoja 'FRANQUETTE' verifique que existe en el archivo", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Btn_Examinar_Click(object sender, EventArgs e)
        {
            OpenFileDialog newdialog = new OpenFileDialog();
            newdialog.Filter = "Excel files(*.xlsx)|*.xls";
            if (newdialog.ShowDialog() == DialogResult.OK)
            {
                TxtFile.Text = newdialog.FileName;
                importarDatosSERR(TxtFile.Text);
                importarDatosCHANDLER(TxtFile.Text);
                importarDatosHOWARD(TxtFile.Text);
                importarDatosFRANQUETTE(TxtFile.Text);
                importarDatosHARTLEY(TxtFile.Text);
                importarDatosSEMILLA(TxtFile.Text);
                importarDatosSUNDLAND(TxtFile.Text);
                importarDatosTULARE(TxtFile.Text);
                importarDatosVINA(TxtFile.Text);
                importarDatosVINA_MEJ(TxtFile.Text);

                if (dataGridView1.Rows.Count > 0)
                {
                    btn_ingresar.Visible = true;
                    Btn_Examinar.Visible = false;
                }
            }
        }

        private void btn_ingresar_Click(object sender, EventArgs e)
        {
            Eliminar_Por_NCC();
            for(int i = 1;i<dataGridView1.ColumnCount;i++)
            {        
               Agregar_Por_NCC(dataGridView1.Columns[i].Name.ToString(),"2",Convert.ToSingle(dataGridView1.Rows[0].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView1.Rows[1].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView1.Rows[2].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView1.Rows[3].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView1.Rows[4].Cells[i].Value.ToString()),Convert.ToSingle(dataGridView1.Rows[5].Cells[i].Value.ToString()));
               Agregar_Por_NCC(dataGridView2.Columns[i].Name.ToString(),"5",Convert.ToSingle(dataGridView2.Rows[0].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView2.Rows[1].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView2.Rows[2].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView2.Rows[3].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView2.Rows[4].Cells[i].Value.ToString()),Convert.ToSingle(dataGridView2.Rows[5].Cells[i].Value.ToString()));
               Agregar_Por_NCC(dataGridView3.Columns[i].Name.ToString(),"9",Convert.ToSingle(dataGridView3.Rows[0].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView3.Rows[1].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView3.Rows[2].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView3.Rows[3].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView3.Rows[4].Cells[i].Value.ToString()),Convert.ToSingle(dataGridView3.Rows[5].Cells[i].Value.ToString()));
               Agregar_Por_NCC(dataGridView4.Columns[i].Name.ToString(),"10",Convert.ToSingle(dataGridView4.Rows[0].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView4.Rows[1].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView4.Rows[2].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView4.Rows[3].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView4.Rows[4].Cells[i].Value.ToString()),Convert.ToSingle(dataGridView4.Rows[5].Cells[i].Value.ToString()));
               Agregar_Por_NCC(dataGridView5.Columns[i].Name.ToString(),"6",Convert.ToSingle(dataGridView5.Rows[0].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView5.Rows[1].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView5.Rows[2].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView5.Rows[3].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView5.Rows[4].Cells[i].Value.ToString()),Convert.ToSingle(dataGridView5.Rows[5].Cells[i].Value.ToString()));
               Agregar_Por_NCC(dataGridView6.Columns[i].Name.ToString(),"4",Convert.ToSingle(dataGridView6.Rows[0].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView6.Rows[1].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView6.Rows[2].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView6.Rows[3].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView6.Rows[4].Cells[i].Value.ToString()),Convert.ToSingle(dataGridView6.Rows[5].Cells[i].Value.ToString()));
               Agregar_Por_NCC(dataGridView7.Columns[i].Name.ToString(),"1",Convert.ToSingle(dataGridView7.Rows[0].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView7.Rows[1].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView7.Rows[2].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView7.Rows[3].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView7.Rows[4].Cells[i].Value.ToString()),Convert.ToSingle(dataGridView7.Rows[5].Cells[i].Value.ToString()));
               Agregar_Por_NCC(dataGridView8.Columns[i].Name.ToString(),"3",Convert.ToSingle(dataGridView8.Rows[0].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView8.Rows[1].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView8.Rows[2].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView8.Rows[3].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView8.Rows[4].Cells[i].Value.ToString()),Convert.ToSingle(dataGridView8.Rows[5].Cells[i].Value.ToString()));
               Agregar_Por_NCC(dataGridView9.Columns[i].Name.ToString(),"15",Convert.ToSingle(dataGridView9.Rows[0].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView9.Rows[1].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView9.Rows[2].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView9.Rows[3].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView9.Rows[4].Cells[i].Value.ToString()),Convert.ToSingle(dataGridView9.Rows[5].Cells[i].Value.ToString()));
               Agregar_Por_NCC(dataGridView10.Columns[i].Name.ToString(),"8",Convert.ToSingle(dataGridView10.Rows[0].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView10.Rows[1].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView10.Rows[2].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView10.Rows[3].Cells[i].Value.ToString()), Convert.ToSingle(dataGridView10.Rows[4].Cells[i].Value.ToString()),Convert.ToSingle(dataGridView10.Rows[5].Cells[i].Value.ToString()));
            }
            MessageBox.Show("Informacion ingresada correctamente", "ANAKENA",MessageBoxButtons.OK,MessageBoxIcon.Information);
        }
        public void Agregar_Por_NCC(string Producto, string variedad, float exportacion,float Pin_Sortex,float Descarte,float Vanas,float blanquear,float Mermas)
        {
            try
            {
                // linea de comando de sql
                SqlCommand cmd = new SqlCommand("spAgregar_Por_NCC", cn.getConexion());
                // adhesion de parametros 
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@Producto", SqlDbType.Text);
                cmd.Parameters.Add("@Cod_Variedad", SqlDbType.Char,2);
                cmd.Parameters.Add("@Exportacion", SqlDbType.Float);
                cmd.Parameters.Add("@Pin_Sortex", SqlDbType.Float);
                cmd.Parameters.Add("@Descarte", SqlDbType.Float);
                cmd.Parameters.Add("@Vanas", SqlDbType.Float);
                cmd.Parameters.Add("@Blanquear", SqlDbType.Float);
                cmd.Parameters.Add("@Merma", SqlDbType.Float);
                cmd.Parameters.Add("@msg", SqlDbType.VarChar, 100);

                // ingreso de parametros
                cmd.Parameters["@Producto"].Value = Producto;
                cmd.Parameters["@Cod_Variedad"].Value = variedad;
                cmd.Parameters["@Exportacion"].Value = exportacion;
                cmd.Parameters["@Pin_Sortex"].Value = Pin_Sortex;
                cmd.Parameters["@Descarte"].Value = Descarte;
                cmd.Parameters["@Vanas"].Value = Vanas;
                cmd.Parameters["@Blanquear"].Value = blanquear;
                cmd.Parameters["@Merma"].Value = Mermas;
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
        public void Eliminar_Por_NCC()
        {
            try
            {
                // linea de comando de sql
                SqlCommand cmd = new SqlCommand("spDeletePor_Var", cn.getConexion());
                // adhesion de parametros 
                cmd.CommandType = CommandType.StoredProcedure;          
                cmd.Parameters.Add("@msg", SqlDbType.VarChar, 100);
                // ingreso de parametros
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
