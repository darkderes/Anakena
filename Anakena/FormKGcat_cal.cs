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
    public partial class FormKGcat_cal : Form
    {
        conexion cn = new conexion();
        public FormKGcat_cal()
        {
            InitializeComponent();
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
        public void importarDatosDESCTE_DESP(string file)
        {
            try
            {
                OleDbCommand command = new OleDbCommand();
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + file + "';Extended Properties=Excel 8.0;";
                command = new OleDbCommand("select * from [DCTE DESPELONADO$]", conn);
                DataSet myds = new DataSet();
                OleDbDataAdapter adap = new OleDbDataAdapter(command);
                adap.Fill(myds);
                dataGridView11.DataSource = myds.Tables[0];
            }
            catch
            {
                MessageBox.Show("Error al tratar cargar Hoja 'FRANQUETTE' verifique que existe en el archivo", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void importarDatosDESCTE_QUEB(string file)
        {
            try
            {
                OleDbCommand command = new OleDbCommand();
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + file + "';Extended Properties=Excel 8.0;";
                command = new OleDbCommand("select * from [DCTE QUEBRADO$]", conn);
                DataSet myds = new DataSet();
                OleDbDataAdapter adap = new OleDbDataAdapter(command);
                adap.Fill(myds);
                dataGridView12.DataSource = myds.Tables[0];
            }
            catch
            {
                MessageBox.Show("Error al tratar cargar Hoja 'DCTE QUEBRADO' verifique que existe en el archivo", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                importarDatosDESCTE_DESP(TxtFile.Text);
                importarDatosDESCTE_QUEB(TxtFile.Text);
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
            pBar1.Minimum = 1;
            pBar1.Maximum = (dataGridView1.RowCount * dataGridView1.ColumnCount);
            pBar1.Value = 1;
            pBar1.Step = 1;
            double acum = 0;
            eliminar_KG_CAT_CAL();
            for (int x = 1; x < dataGridView1.Columns.Count-1; x++)
            {
                for (int i = 0; i < dataGridView1.RowCount-1; i++)
                {
                    pBar1.Visible = true;
                    pBar1.PerformStep();
                    Ingresar_KG_CAT_CAL(traerId_Calibre(dataGridView1.Rows[i].Cells[0].Value.ToString()), traerId_Categoria(dataGridView1.Columns[x].Name), Convert.ToSingle(dataGridView1.Rows[i].Cells[x].Value.ToString()), "2");
                    Ingresar_KG_CAT_CAL(traerId_Calibre(dataGridView2.Rows[i].Cells[0].Value.ToString()), traerId_Categoria(dataGridView2.Columns[x].Name), Convert.ToSingle(dataGridView2.Rows[i].Cells[x].Value.ToString()), "5");
                    Ingresar_KG_CAT_CAL(traerId_Calibre(dataGridView3.Rows[i].Cells[0].Value.ToString()), traerId_Categoria(dataGridView3.Columns[x].Name), Convert.ToSingle(dataGridView3.Rows[i].Cells[x].Value.ToString()), "9");
                    Ingresar_KG_CAT_CAL(traerId_Calibre(dataGridView4.Rows[i].Cells[0].Value.ToString()), traerId_Categoria(dataGridView4.Columns[x].Name), Convert.ToSingle(dataGridView4.Rows[i].Cells[x].Value.ToString()), "10");
                    Ingresar_KG_CAT_CAL(traerId_Calibre(dataGridView5.Rows[i].Cells[0].Value.ToString()), traerId_Categoria(dataGridView5.Columns[x].Name), Convert.ToSingle(dataGridView5.Rows[i].Cells[x].Value.ToString()), "6");
                    Ingresar_KG_CAT_CAL(traerId_Calibre(dataGridView6.Rows[i].Cells[0].Value.ToString()), traerId_Categoria(dataGridView6.Columns[x].Name), Convert.ToSingle(dataGridView6.Rows[i].Cells[x].Value.ToString()), "4");
                    Ingresar_KG_CAT_CAL(traerId_Calibre(dataGridView7.Rows[i].Cells[0].Value.ToString()), traerId_Categoria(dataGridView7.Columns[x].Name), Convert.ToSingle(dataGridView7.Rows[i].Cells[x].Value.ToString()), "1");
                    Ingresar_KG_CAT_CAL(traerId_Calibre(dataGridView8.Rows[i].Cells[0].Value.ToString()), traerId_Categoria(dataGridView8.Columns[x].Name), Convert.ToSingle(dataGridView8.Rows[i].Cells[x].Value.ToString()), "3");
                    Ingresar_KG_CAT_CAL(traerId_Calibre(dataGridView9.Rows[i].Cells[0].Value.ToString()), traerId_Categoria(dataGridView9.Columns[x].Name), Convert.ToSingle(dataGridView9.Rows[i].Cells[x].Value.ToString()), "15");
                    Ingresar_KG_CAT_CAL(traerId_Calibre(dataGridView10.Rows[i].Cells[0].Value.ToString()), traerId_Categoria(dataGridView10.Columns[x].Name), Convert.ToSingle(dataGridView10.Rows[i].Cells[x].Value.ToString()), "8");
                    Ingresar_KG_CAT_CAL(traerId_Calibre(dataGridView11.Rows[i].Cells[0].Value.ToString()), traerId_Categoria(dataGridView11.Columns[x].Name), Convert.ToSingle(dataGridView11.Rows[i].Cells[x].Value.ToString()), "16");
                    Ingresar_KG_CAT_CAL(traerId_Calibre(dataGridView12.Rows[i].Cells[0].Value.ToString()), traerId_Categoria(dataGridView12.Columns[x].Name), Convert.ToSingle(dataGridView12.Rows[i].Cells[x].Value.ToString()), "14");
                    acum = acum + (Convert.ToSingle(dataGridView1.Rows[i].Cells[x].Value.ToString()) + Convert.ToSingle(dataGridView2.Rows[i].Cells[x].Value.ToString()) + Convert.ToSingle(dataGridView3.Rows[i].Cells[x].Value.ToString()) + Convert.ToSingle(dataGridView4.Rows[i].Cells[x].Value.ToString()) + Convert.ToSingle(dataGridView5.Rows[i].Cells[x].Value.ToString()) + Convert.ToSingle(dataGridView6.Rows[i].Cells[x].Value.ToString()) + Convert.ToSingle(dataGridView7.Rows[i].Cells[x].Value.ToString()) + Convert.ToSingle(dataGridView8.Rows[i].Cells[x].Value.ToString()) + Convert.ToSingle(dataGridView9.Rows[i].Cells[x].Value.ToString()) + Convert.ToSingle(dataGridView10.Rows[i].Cells[x].Value.ToString()) + Convert.ToSingle(dataGridView11.Rows[i].Cells[x].Value.ToString()) + Convert.ToSingle(dataGridView12.Rows[i].Cells[x].Value.ToString()));
                }
            }
            pBar1.Visible = false;
            MessageBox.Show("KG_CAT_CAL ingresada correctamente con "+acum.ToString("##,##.##") +" Kg", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        }
        public string traerId_Calibre(string calibre)
        {
            string id = "";
            SqlCommand cmd = new SqlCommand("spTraerIdCalibre", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@calibre", SqlDbType.VarChar, 25);
            cmd.Parameters["@calibre"].Value = calibre;
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
        private void eliminar_KG_CAT_CAL()
        {
            try
            {
                // linea de comando de sql
                SqlCommand cmd = new SqlCommand("spDeleteKG_Cat_Cal", cn.getConexion());

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
            finally
            {


            }

        }
        public void Ingresar_KG_CAT_CAL( string calibre,string categoria, float monto, string variedad)
        {
            try
            {
                // linea de comando de sql
                SqlCommand cmd = new SqlCommand("spAgregar_KG_CAT_CAL", cn.getConexion());

                // adhrsion de parametros 
                cmd.CommandType = CommandType.StoredProcedure;
               
                cmd.Parameters.Add("@Cod_Calibre", SqlDbType.Char, 2);
                cmd.Parameters.Add("@Cod_Categoria", SqlDbType.Char,2);
                cmd.Parameters.Add("@Monto", SqlDbType.Float);
                cmd.Parameters.Add("@Cod_Variedad", SqlDbType.Char, 2);


                cmd.Parameters.Add("@msg", SqlDbType.VarChar, 100);

                // ingreso de parametros
               
                cmd.Parameters["@Cod_Calibre"].Value = calibre;
                cmd.Parameters["@Cod_Categoria"].Value = categoria;
                cmd.Parameters["@Monto"].Value = monto;
                cmd.Parameters["@Cod_Variedad"].Value = variedad;


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
        private void FormKGcat_cal_Load(object sender, EventArgs e)
        {

        }
    }
}
