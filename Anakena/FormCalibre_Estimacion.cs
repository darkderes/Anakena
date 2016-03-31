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
    public partial class FormCalibre_Estimacion : Form
    {
        conexion cn = new conexion();
        int x = -2;
        public FormCalibre_Estimacion()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
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
                MessageBox.Show("Error al tratar cargar Hoja 'CHANDLER' verifique que existe en el archivo","Anakena",MessageBoxButtons.OK,MessageBoxIcon.Error);
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
                MessageBox.Show("Error al tratar cargar Hoja 'HOWARDS' verifique que existe en el archivo","Anakena",MessageBoxButtons.OK,MessageBoxIcon.Error);
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
        private void FormCalibre_Estimacion_Load(object sender, EventArgs e)
        {
            CmbFechas();
            traer_serr();
            traer_chandler();
            traer_howard();
            traer_tulare();
            traer_sundland();
            traer_hartley();
            traer_semilla();
            traer_vina();
            traer_vina_mej();
            traer_franquette();
            traer_despelonado();
            traer_quebrado();
        }
        public void traer_serr()
        {
            SqlCommand cmd = new SqlCommand("spEstimacionCalibreporfecha", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Int);
            cmd.Parameters["@variedad"].Value = 2;
            cmd.Parameters.Add("@fecha", SqlDbType.Char, 10);
            cmd.Parameters["@fecha"].Value = Cmb_Fechas.Text ;
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView1.DataSource = myds.Tables[0];
            cn.Cerrar();
        }
        public void traer_chandler()
        {
            SqlCommand cmd = new SqlCommand("spEstimacionCalibreporfecha", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Int);
            cmd.Parameters["@variedad"].Value = 5;
            cmd.Parameters.Add("@fecha", SqlDbType.Char, 10);
            cmd.Parameters["@fecha"].Value = Cmb_Fechas.Text;
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView2.DataSource = myds.Tables[0];
            cn.Cerrar();
        }
        public void traer_howard()
        {
            SqlCommand cmd = new SqlCommand("spEstimacionCalibreporfecha", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Int);
            cmd.Parameters["@variedad"].Value = 9;
            cmd.Parameters.Add("@fecha", SqlDbType.Char, 10);
            cmd.Parameters["@fecha"].Value = Cmb_Fechas.Text;
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView3.DataSource = myds.Tables[0];
            cn.Cerrar();
        }

        public void traer_tulare()
        {
            SqlCommand cmd = new SqlCommand("spEstimacionCalibreporfecha", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Int);
            cmd.Parameters["@variedad"].Value = 10;
            cmd.Parameters.Add("@fecha", SqlDbType.Char, 10);
            cmd.Parameters["@fecha"].Value = Cmb_Fechas.Text;
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView4.DataSource = myds.Tables[0];
            cn.Cerrar();
        }
        public void traer_sundland()
        {
            SqlCommand cmd = new SqlCommand("spEstimacionCalibreporfecha", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Int);
            cmd.Parameters["@variedad"].Value = 6;
            cmd.Parameters.Add("@fecha", SqlDbType.Char, 10);
            cmd.Parameters["@fecha"].Value = Cmb_Fechas.Text;
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView5.DataSource = myds.Tables[0];
            cn.Cerrar();
        }
        public void traer_hartley()
        {
            SqlCommand cmd = new SqlCommand("spEstimacionCalibreporfecha", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Int);
            cmd.Parameters["@variedad"].Value = 4;
            cmd.Parameters.Add("@fecha", SqlDbType.Char, 10);
            cmd.Parameters["@fecha"].Value = Cmb_Fechas.Text;
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView6.DataSource = myds.Tables[0];
            cn.Cerrar();
        }
        public void traer_semilla()
        {
            SqlCommand cmd = new SqlCommand("spEstimacionCalibreporfecha", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Int);
            cmd.Parameters["@variedad"].Value = 1;
            cmd.Parameters.Add("@fecha", SqlDbType.Char, 10);
            cmd.Parameters["@fecha"].Value = Cmb_Fechas.Text;
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView7.DataSource = myds.Tables[0];
            cn.Cerrar();
        }
        public void traer_vina()
        {
            SqlCommand cmd = new SqlCommand("spEstimacionCalibreporfecha", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Int);
            cmd.Parameters["@variedad"].Value = 3;
            cmd.Parameters.Add("@fecha", SqlDbType.Char, 10);
            cmd.Parameters["@fecha"].Value = Cmb_Fechas.Text;
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView8.DataSource = myds.Tables[0];
            cn.Cerrar();
        }
        public void traer_vina_mej()
        {
            SqlCommand cmd = new SqlCommand("spEstimacionCalibreporfecha", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Int);
            cmd.Parameters["@variedad"].Value = 15;
            cmd.Parameters.Add("@fecha", SqlDbType.Char, 10);
            cmd.Parameters["@fecha"].Value = Cmb_Fechas.Text;
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView9.DataSource = myds.Tables[0];
            cn.Cerrar();
        }
        public void traer_franquette()
        {
            SqlCommand cmd = new SqlCommand("spEstimacionCalibreporfecha", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Int);
            cmd.Parameters["@variedad"].Value = 8;
            cmd.Parameters.Add("@fecha", SqlDbType.Char, 10);
            cmd.Parameters["@fecha"].Value = Cmb_Fechas.Text;
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView10.DataSource = myds.Tables[0];
            cn.Cerrar();
        }
        public void traer_despelonado()
        {
            SqlCommand cmd = new SqlCommand("spEstimacionCalibreporfecha", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Int);
            cmd.Parameters["@variedad"].Value = 16;
            cmd.Parameters.Add("@fecha", SqlDbType.Char, 10);
            cmd.Parameters["@fecha"].Value = Cmb_Fechas.Text;
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView11.DataSource = myds.Tables[0];
            cn.Cerrar();
        }
        public void traer_quebrado()
        {
            SqlCommand cmd = new SqlCommand("spEstimacionCalibreporfecha", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Int);
            cmd.Parameters["@variedad"].Value = 14;
            cmd.Parameters.Add("@fecha", SqlDbType.Char, 10);
            cmd.Parameters["@fecha"].Value = Cmb_Fechas.Text;
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView12.DataSource = myds.Tables[0];
            cn.Cerrar();
        }
        public void CmbFechas()
        {
            try
            {
                cn.Abrir();
                SqlCommand command = new SqlCommand("spTraerFechaEstimacionCalibre", cn.getConexion());
                SqlDataAdapter myAdapter = new SqlDataAdapter(command);
                command.CommandType = CommandType.StoredProcedure;
                DataSet myds = new DataSet();
                myAdapter.Fill(myds, "Estimacion");
                Cmb_Fechas.Refresh();
                Cmb_Fechas.DataSource = myds.Tables["Estimacion"].DefaultView;
                Cmb_Fechas.DisplayMember = "fecha";
                Cmb_Fechas.ValueMember = "fecha";
            }
            catch (Exception e)
            {
                MessageBox.Show("Ocurrio un problema al cargar fechas :" + e, "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cn.Cerrar();
            }
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {

          // e.CellStyle.Format = "P02";
        }

        private void btn_ingresar_Click(object sender, EventArgs e)
        {   pBar1.Minimum = 1;
            pBar1.Maximum = (dataGridView1.RowCount * dataGridView1.ColumnCount);
            pBar1.Value = 1;
            pBar1.Step = 1;
            eliminar_estimacion();
         
            for (int x = 1; x < dataGridView1.Columns.Count; x++)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
               {
                  pBar1.Visible = true;
                  pBar1.PerformStep();
                  ingresarEstimacionCalibre(dataGridView1.Columns[x].Name,traerId_Calibre(dataGridView1.Rows[i].Cells[0].Value.ToString()),Convert.ToSingle(dataGridView1.Rows[i].Cells[x].Value.ToString()),"2");
                  ingresarEstimacionCalibre(dataGridView2.Columns[x].Name,traerId_Calibre(dataGridView2.Rows[i].Cells[0].Value.ToString()),Convert.ToSingle(dataGridView2.Rows[i].Cells[x].Value.ToString()),"5");
                  ingresarEstimacionCalibre(dataGridView3.Columns[x].Name,traerId_Calibre(dataGridView3.Rows[i].Cells[0].Value.ToString()),Convert.ToSingle(dataGridView3.Rows[i].Cells[x].Value.ToString()),"9");
                  ingresarEstimacionCalibre(dataGridView4.Columns[x].Name,traerId_Calibre(dataGridView4.Rows[i].Cells[0].Value.ToString()),Convert.ToSingle(dataGridView4.Rows[i].Cells[x].Value.ToString()),"10");
                  ingresarEstimacionCalibre(dataGridView5.Columns[x].Name,traerId_Calibre(dataGridView5.Rows[i].Cells[0].Value.ToString()),Convert.ToSingle(dataGridView5.Rows[i].Cells[x].Value.ToString()),"6");
                  ingresarEstimacionCalibre(dataGridView6.Columns[x].Name,traerId_Calibre(dataGridView6.Rows[i].Cells[0].Value.ToString()),Convert.ToSingle(dataGridView6.Rows[i].Cells[x].Value.ToString()),"4");
                  ingresarEstimacionCalibre(dataGridView7.Columns[x].Name,traerId_Calibre(dataGridView7.Rows[i].Cells[0].Value.ToString()),Convert.ToSingle(dataGridView7.Rows[i].Cells[x].Value.ToString()),"1");
                  ingresarEstimacionCalibre(dataGridView8.Columns[x].Name,traerId_Calibre(dataGridView8.Rows[i].Cells[0].Value.ToString()),Convert.ToSingle(dataGridView8.Rows[i].Cells[x].Value.ToString()),"3");
                  ingresarEstimacionCalibre(dataGridView9.Columns[x].Name,traerId_Calibre(dataGridView9.Rows[i].Cells[0].Value.ToString()),Convert.ToSingle(dataGridView9.Rows[i].Cells[x].Value.ToString()),"15");
                  ingresarEstimacionCalibre(dataGridView10.Columns[x].Name,traerId_Calibre(dataGridView10.Rows[i].Cells[0].Value.ToString()),Convert.ToSingle(dataGridView10.Rows[i].Cells[x].Value.ToString()),"8");
                  ingresarEstimacionCalibre(dataGridView11.Columns[x].Name,traerId_Calibre(dataGridView11.Rows[i].Cells[0].Value.ToString()),Convert.ToSingle(dataGridView11.Rows[i].Cells[x].Value.ToString()),"16");
                  ingresarEstimacionCalibre(dataGridView12.Columns[x].Name,traerId_Calibre(dataGridView12.Rows[i].Cells[0].Value.ToString()),Convert.ToSingle(dataGridView12.Rows[i].Cells[x].Value.ToString()),"14");
                }
            }
            pBar1.Visible = false;
            MessageBox.Show("Estimacion ingresada correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        }
        public string traerId_Calibre(string calibre)
        {
            string id = "";
            SqlCommand cmd = new SqlCommand("spTraerIdCalibre", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@calibre", SqlDbType.VarChar,25);
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
        private void eliminar_estimacion()
        {
            try
            {
                // linea de comando de sql
                SqlCommand cmd = new SqlCommand("spDeleteEstimacion_Calibre", cn.getConexion());

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
        public void ingresarEstimacionCalibre(string codigo,string calibre,float porcentaje,string variedad)
        {
            try
            {
                // linea de comando de sql
                SqlCommand cmd = new SqlCommand("spAgregaEstimacionCalibre", cn.getConexion());

                // adhrsion de parametros 
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@Cod_Productor", SqlDbType.Char, 10);
                cmd.Parameters.Add("@Cod_Calibre", SqlDbType.Char, 2);
                cmd.Parameters.Add("@porcentaje", SqlDbType.Float);
                cmd.Parameters.Add("@Cod_Variedad", SqlDbType.Char, 2);


                cmd.Parameters.Add("@msg", SqlDbType.VarChar, 100);

                // ingreso de parametros
                cmd.Parameters["@Cod_Productor"].Value = codigo;
                cmd.Parameters["@Cod_Calibre"].Value = calibre;
                cmd.Parameters["@porcentaje"].Value = porcentaje;
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

        private void dataGridView2_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            //e.CellStyle.Format = "P02";
        }

        private void dataGridView3_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            //e.CellStyle.Format = "P02";
        }

        private void dataGridView4_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            //e.CellStyle.Format = "P02";
        }

        private void dataGridView5_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            //e.CellStyle.Format = "P02";
        }

        private void dataGridView6_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            //e.CellStyle.Format = "P02";
        }

        private void dataGridView7_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            //e.CellStyle.Format = "P02";
        }

        private void dataGridView8_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            //e.CellStyle.Format = "P02";
        }

        private void dataGridView9_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            //e.CellStyle.Format = "P02";
        }

        private void dataGridView10_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            //e.CellStyle.Format = "P02";
        }

        private void dataGridView11_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            //e.CellStyle.Format = "P02";
        }

        private void dataGridView12_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            //e.CellStyle.Format = "P02";
        }

        private void Cmb_Fechas_SelectedIndexChanged(object sender, EventArgs e)
        {
            x++;
           if(x > 0)
            {
                Cursor = Cursors.WaitCursor;
                traer_serr();
                traer_chandler();
                traer_howard();
                traer_tulare();
                traer_sundland();
                traer_hartley();
                traer_semilla();
                traer_vina();
                traer_vina_mej();
                traer_franquette();
                traer_despelonado();
                traer_quebrado();
                Cursor = Cursors.Default;
                MessageBox.Show("Datos cargados con fecha " + Cmb_Fechas.Text, "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
