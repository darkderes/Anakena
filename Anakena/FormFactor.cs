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
    public partial class FormFactor : Form
    {
        conexion cn = new conexion();
        DataSet myds = new DataSet();
        DataSet myds_serr = new DataSet();
        public FormFactor()
        {
            InitializeComponent();
        }
        public void CmbProductor()
        {
            try
            {
                cn.Abrir();
                SqlCommand command = new SqlCommand("spDistintosProductores", cn.getConexion());
                SqlDataAdapter myAdapter = new SqlDataAdapter(command);
                command.CommandType = CommandType.StoredProcedure;
                DataSet myds = new DataSet();
                myAdapter.Fill(myds, "Productores");
                cmb_productor.Refresh();
                cmb_productor.DataSource = myds.Tables["Productores"].DefaultView;
                cmb_productor.DisplayMember = "Nom_Productor";
                cmb_productor.ValueMember = "Cod_Productor";
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
        public void importarDatosSERR(string file)
        {

            OleDbCommand command = new OleDbCommand();
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + file + "';Extended Properties=Excel 8.0;";
            command = new OleDbCommand("select * from [SERR$]", conn);
            OleDbDataAdapter adap = new OleDbDataAdapter(command);
            adap.Fill(myds);
            dataGridView1.DataSource = myds.Tables[0];

        }
        public void importarDatosCHANDLER(string file)
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
        public void importarDatosHOWARD(string file)
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
        public void importarDatosTULARE(string file)
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
        public void importarDatosSUNDLAND(string file)
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
        public void importarDatosHARTLEY(string file)
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
        public void importarDatosSEMILLA(string file)
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
        public void importarDatosVINA(string file)
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
        public void importarDatosVINA_MEJ(string file)
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
        public void importarDatosFRANQUETTE(string file)
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
        public void importarDatosDESCTE_DESP(string file)
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
        public void importarDatosDESCTE_QUEB(string file)
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
        private void Btn_Examinar_Click(object sender, EventArgs e)
        {
            EliminarFactor();
            pBar1.Minimum = 1;
            pBar1.Maximum = dataGridView1.RowCount;
            pBar1.Value = 1;
            pBar1.Step = 1;
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                try
                {
                    string productor = dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                    string variedad = dataGridView1.Rows[i].Cells["Variedad"].Value.ToString();
                    double estimacion = Convert.ToDouble(dataGridView1.Rows[i].Cells["Porcentaje_estimacion"].Value.ToString());
                    double realidad = Convert.ToDouble(dataGridView1.Rows[i].Cells["Porcentaje_Realidad"].Value.ToString());
                    ingresarFactor(productor, variedad, estimacion, realidad);
                    ingresarFactor(dataGridView2.Rows[i].Cells["productor"].Value.ToString(), dataGridView2.Rows[i].Cells["Variedad"].Value.ToString(), Convert.ToDouble(dataGridView2.Rows[i].Cells["Porcentaje_estimacion"].Value.ToString()), Convert.ToDouble(dataGridView2.Rows[i].Cells["Porcentaje_Realidad"].Value.ToString()));
                    ingresarFactor(dataGridView3.Rows[i].Cells["productor"].Value.ToString(), dataGridView3.Rows[i].Cells["Variedad"].Value.ToString(), Convert.ToDouble(dataGridView3.Rows[i].Cells["Porcentaje_estimacion"].Value.ToString()), Convert.ToDouble(dataGridView3.Rows[i].Cells["Porcentaje_Realidad"].Value.ToString()));
                    ingresarFactor(dataGridView4.Rows[i].Cells["productor"].Value.ToString(), dataGridView4.Rows[i].Cells["Variedad"].Value.ToString(), Convert.ToDouble(dataGridView4.Rows[i].Cells["Porcentaje_estimacion"].Value.ToString()), Convert.ToDouble(dataGridView4.Rows[i].Cells["Porcentaje_Realidad"].Value.ToString()));
                    ingresarFactor(dataGridView5.Rows[i].Cells["productor"].Value.ToString(), dataGridView5.Rows[i].Cells["Variedad"].Value.ToString(), Convert.ToDouble(dataGridView5.Rows[i].Cells["Porcentaje_estimacion"].Value.ToString()), Convert.ToDouble(dataGridView5.Rows[i].Cells["Porcentaje_Realidad"].Value.ToString()));
                    ingresarFactor(dataGridView6.Rows[i].Cells["productor"].Value.ToString(), dataGridView6.Rows[i].Cells["Variedad"].Value.ToString(), Convert.ToDouble(dataGridView6.Rows[i].Cells["Porcentaje_estimacion"].Value.ToString()), Convert.ToDouble(dataGridView6.Rows[i].Cells["Porcentaje_Realidad"].Value.ToString()));
                    ingresarFactor(dataGridView7.Rows[i].Cells["productor"].Value.ToString(), dataGridView7.Rows[i].Cells["Variedad"].Value.ToString(), Convert.ToDouble(dataGridView7.Rows[i].Cells["Porcentaje_estimacion"].Value.ToString()), Convert.ToDouble(dataGridView7.Rows[i].Cells["Porcentaje_Realidad"].Value.ToString()));
                    ingresarFactor(dataGridView8.Rows[i].Cells["productor"].Value.ToString(), dataGridView8.Rows[i].Cells["Variedad"].Value.ToString(), Convert.ToDouble(dataGridView8.Rows[i].Cells["Porcentaje_estimacion"].Value.ToString()), Convert.ToDouble(dataGridView8.Rows[i].Cells["Porcentaje_Realidad"].Value.ToString()));
                    ingresarFactor(dataGridView9.Rows[i].Cells["productor"].Value.ToString(), dataGridView9.Rows[i].Cells["Variedad"].Value.ToString(), Convert.ToDouble(dataGridView9.Rows[i].Cells["Porcentaje_estimacion"].Value.ToString()), Convert.ToDouble(dataGridView9.Rows[i].Cells["Porcentaje_Realidad"].Value.ToString()));
                    ingresarFactor(dataGridView10.Rows[i].Cells["productor"].Value.ToString(), dataGridView10.Rows[i].Cells["Variedad"].Value.ToString(), Convert.ToDouble(dataGridView10.Rows[i].Cells["Porcentaje_estimacion"].Value.ToString()), Convert.ToDouble(dataGridView10.Rows[i].Cells["Porcentaje_Realidad"].Value.ToString()));
                    ingresarFactor(dataGridView11.Rows[i].Cells["productor"].Value.ToString(), dataGridView11.Rows[i].Cells["Variedad"].Value.ToString(), Convert.ToDouble(dataGridView11.Rows[i].Cells["Porcentaje_estimacion"].Value.ToString()), Convert.ToDouble(dataGridView11.Rows[i].Cells["Porcentaje_Realidad"].Value.ToString()));
                    ingresarFactor(dataGridView12.Rows[i].Cells["productor"].Value.ToString(), dataGridView12.Rows[i].Cells["Variedad"].Value.ToString(), Convert.ToDouble(dataGridView12.Rows[i].Cells["Porcentaje_estimacion"].Value.ToString()), Convert.ToDouble(dataGridView12.Rows[i].Cells["Porcentaje_Realidad"].Value.ToString()));

                    pBar1.PerformStep();
                }
                catch { MessageBox.Show("Error al tratar de ingresar los datos", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                finally { }

            }
            MessageBox.Show("Factor ingresados correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        }
        public void ingresarFactor(string codigo,string variedad,double porcentaje_est,double porcentaje_real)
        {
            try
            {
                // linea de comando de sql
                SqlCommand cmd = new SqlCommand("spAgregaFactores", cn.getConexion());

                // adhrsion de parametros 
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@Cod_Productor", SqlDbType.Char, 10);
                cmd.Parameters.Add("@Cod_Variedad", SqlDbType.Char, 2);
                cmd.Parameters.Add("@Porcentaje_estimacion", SqlDbType.Float);
                cmd.Parameters.Add("@Porcentaje_real", SqlDbType.Float);


                cmd.Parameters.Add("@msg", SqlDbType.VarChar, 100);

                // ingreso de parametros
                cmd.Parameters["@Cod_Productor"].Value = codigo;
               
                cmd.Parameters["@Cod_Variedad"].Value = variedad;
                cmd.Parameters["@Porcentaje_estimacion"].Value = porcentaje_est;
                cmd.Parameters["@Porcentaje_real"].Value = porcentaje_real;


                cmd.Parameters["@msg"].Value = 1;
                //abre coneccion a base de datos
                cn.Abrir();
                cmd.ExecuteNonQuery();
                string sqlMensaje = cmd.Parameters["@msg"].Value.ToString();

                //cierre de coneccion a base de datos

            }
            catch (Exception ex)
            {
                //mensaje de error SQL
                MessageBox.Show("Error SQL " + ex, "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cn.Cerrar();
            }
        }


        public void ModificarFactor(string codigo, double porcentaje_est, double porcentaje_real)
        {
            try
            {
                // linea de comando de sql
                SqlCommand cmd = new SqlCommand("Update_Factor", cn.getConexion());

                // adhrsion de parametros 
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@id", SqlDbType.Char, 10);
                cmd.Parameters.Add("@estimacion", SqlDbType.Float);
                cmd.Parameters.Add("@real", SqlDbType.Float);


                cmd.Parameters.Add("@msg", SqlDbType.VarChar, 100);
                // ingreso de parametros
                cmd.Parameters["@id"].Value = codigo;
                cmd.Parameters["@estimacion"].Value = porcentaje_est;
                cmd.Parameters["@real"].Value = porcentaje_real;


                cmd.Parameters["@msg"].Value = 1;
                //abre coneccion a base de datos
                cn.Abrir();
                cmd.ExecuteNonQuery();
                string sqlMensaje = cmd.Parameters["@msg"].Value.ToString();

                //cierre de coneccion a base de datos

            }
            catch (Exception ex)
            {
                //mensaje de error SQL
                MessageBox.Show("Error SQL " + ex, "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cn.Cerrar();
            }
        }

        public void EliminarFactor()
        {
            try
            {
                // linea de comando de sql
                SqlCommand cmd = new SqlCommand("spBorrar_Factor", cn.getConexion());
                // adhrsion de parametros 
                cmd.CommandType = CommandType.StoredProcedure;
                cn.Abrir();
                cmd.ExecuteNonQuery();

                //cierre de coneccion a base de datos

            }
            catch (Exception ex)
            {
                //mensaje de error SQL
                MessageBox.Show("Error SQL " + ex, "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cn.Cerrar();
            }
        }
        public void traerFactor(string variedad)
        {
            try
            {
            SqlCommand cmd = new SqlCommand("spTraer_Factor", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Char, 2);
            cmd.Parameters["@variedad"].Value = variedad;
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
             myds = new DataSet();
            adapter.Fill(myds);
            dataGridView1.DataSource = myds.Tables[0];
            cn.Cerrar();
            }
            catch { }

        }
        public void traerFactor_Chandler(string variedad)
        {
            try
            {
                SqlCommand cmd = new SqlCommand("spTraer_Factor", cn.getConexion());
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@variedad", SqlDbType.Char, 2);
                cmd.Parameters["@variedad"].Value = variedad;
                cn.Abrir();
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                myds = new DataSet();
                adapter.Fill(myds);
                dataGridView2.DataSource = myds.Tables[0];
                cn.Cerrar();
            }
            catch { }

        }
        public void traerFactor_HOWARD(string variedad)
        {
            try
            {
                SqlCommand cmd = new SqlCommand("spTraer_Factor", cn.getConexion());
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@variedad", SqlDbType.Char, 2);
                cmd.Parameters["@variedad"].Value = variedad;
                cn.Abrir();
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                myds = new DataSet();
                adapter.Fill(myds);
                dataGridView3.DataSource = myds.Tables[0];
                cn.Cerrar();
            }
            catch { }

        }
        public void traerFactor_TULARE(string variedad)
        {
            try
            {
                SqlCommand cmd = new SqlCommand("spTraer_Factor", cn.getConexion());
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@variedad", SqlDbType.Char, 2);
                cmd.Parameters["@variedad"].Value = variedad;
                cn.Abrir();
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                myds = new DataSet();
                adapter.Fill(myds);
                dataGridView4.DataSource = myds.Tables[0];
                cn.Cerrar();
            }
            catch { }

        }
        public void traerFactor_SUNDLAND(string variedad)
        {
            try
            {
                SqlCommand cmd = new SqlCommand("spTraer_Factor", cn.getConexion());
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@variedad", SqlDbType.Char, 2);
                cmd.Parameters["@variedad"].Value = variedad;
                cn.Abrir();
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                myds = new DataSet();
                adapter.Fill(myds);
                dataGridView5.DataSource = myds.Tables[0];
                cn.Cerrar();
            }
            catch { }

        }

        public void traerFactor_HARTLEY(string variedad)
        {
            try
            {
                SqlCommand cmd = new SqlCommand("spTraer_Factor", cn.getConexion());
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@variedad", SqlDbType.Char, 2);
                cmd.Parameters["@variedad"].Value = variedad;
                cn.Abrir();
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                myds = new DataSet();
                adapter.Fill(myds);
                dataGridView6.DataSource = myds.Tables[0];
                cn.Cerrar();
            }
            catch { }

        }
        public void traerFactor_SEMILLA(string variedad)
        {
            try
            {
                SqlCommand cmd = new SqlCommand("spTraer_Factor", cn.getConexion());
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@variedad", SqlDbType.Char, 2);
                cmd.Parameters["@variedad"].Value = variedad;
                cn.Abrir();
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                myds = new DataSet();
                adapter.Fill(myds);
                dataGridView7.DataSource = myds.Tables[0];
                cn.Cerrar();
            }
            catch { }

        }
        public void traerFactor_VINA(string variedad)
        {
            try
            {
                SqlCommand cmd = new SqlCommand("spTraer_Factor", cn.getConexion());
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@variedad", SqlDbType.Char, 2);
                cmd.Parameters["@variedad"].Value = variedad;
                cn.Abrir();
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                myds = new DataSet();
                adapter.Fill(myds);
                dataGridView8.DataSource = myds.Tables[0];
                cn.Cerrar();
            }
            catch { }

        }
        public void traerFactor_VINA_MEJORADA(string variedad)
        {
            try
            {
                SqlCommand cmd = new SqlCommand("spTraer_Factor", cn.getConexion());
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@variedad", SqlDbType.Char, 2);
                cmd.Parameters["@variedad"].Value = variedad;
                cn.Abrir();
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                myds = new DataSet();
                adapter.Fill(myds);
                dataGridView9.DataSource = myds.Tables[0];
                cn.Cerrar();
            }
            catch { }

        }
        public void traerFactor_FRANQUETTE(string variedad)
        {
            try
            {
                SqlCommand cmd = new SqlCommand("spTraer_Factor", cn.getConexion());
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@variedad", SqlDbType.Char, 2);
                cmd.Parameters["@variedad"].Value = variedad;
                cn.Abrir();
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                myds = new DataSet();
                adapter.Fill(myds);
                dataGridView10.DataSource = myds.Tables[0];
                cn.Cerrar();
            }
            catch { }

        }
        public void traerFactor_DESPELONADO(string variedad)
        {
            try
            {
                SqlCommand cmd = new SqlCommand("spTraer_Factor", cn.getConexion());
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@variedad", SqlDbType.Char, 2);
                cmd.Parameters["@variedad"].Value = variedad;
                cn.Abrir();
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                myds = new DataSet();
                adapter.Fill(myds);
                dataGridView11.DataSource = myds.Tables[0];
                cn.Cerrar();
            }
            catch { }

        }
        public void traerFactor_QUEBRADO(string variedad)
        {
            try
            {
                SqlCommand cmd = new SqlCommand("spTraer_Factor", cn.getConexion());
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@variedad", SqlDbType.Char, 2);
                cmd.Parameters["@variedad"].Value = variedad;
                cn.Abrir();
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                myds = new DataSet();
                adapter.Fill(myds);
                dataGridView12.DataSource = myds.Tables[0];
                cn.Cerrar();
            }
            catch { }

        }
        public void traerFactor_filtro(string productor)
        {
            try
            {
                SqlCommand cmd = new SqlCommand("spProductor_Factor", cn.getConexion());
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@Productor", SqlDbType.Char,10);
                cmd.Parameters["@Productor"].Value = productor;
                cn.Abrir();
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                myds = new DataSet();
                adapter.Fill(myds);
                dataGridView13.DataSource = myds.Tables[0];
                cn.Cerrar();
            }
            catch { }

        }
        private void btn_ingresar_Click(object sender, EventArgs e)
        {
            FormDialogo s = new FormDialogo();
            s.ShowDialog();
            if(s.radioButton1.Checked == true)
            {
                OpenFileDialog newdialog = new OpenFileDialog();
                newdialog.Filter = "Excel files(*.xlsx)|*.xls";
                if (newdialog.ShowDialog() == DialogResult.OK)
                {
                    TxtFile.Text = newdialog.FileName;
                    importarDatosSERR(TxtFile.Text);
                    importarDatosCHANDLER(TxtFile.Text);
                    importarDatosDESCTE_DESP(TxtFile.Text);
                    importarDatosDESCTE_QUEB(TxtFile.Text);
                    importarDatosFRANQUETTE(TxtFile.Text);
                    importarDatosHARTLEY(TxtFile.Text);
                    importarDatosHOWARD(TxtFile.Text);
                    importarDatosSEMILLA(TxtFile.Text);
                    importarDatosSUNDLAND(TxtFile.Text);
                    importarDatosTULARE(TxtFile.Text);
                    importarDatosVINA(TxtFile.Text);
                    importarDatosVINA_MEJ(TxtFile.Text);

                    if (dataGridView1.Rows.Count > 0)
                    {
                        btn_ingresar.Visible = false;
                        Btn_Examinar.Visible = true;
                    }
                }
            }
            else
            {
                s.groupBox1.Visible = false;
                s.groupBox2.Visible = true;
                s.CmbVariedad();
                s.btn_aceptar.Visible = false;
                s.btn_guardar.Visible = true;
                s.ShowDialog();
                traerFactor("2");
                traerFactor_Chandler("5");
                traerFactor_HOWARD("9");
                traerFactor_TULARE("10");
                traerFactor_SUNDLAND("6");
                traerFactor_HARTLEY("4");
                traerFactor_SEMILLA("1");
                traerFactor_VINA("3");
                traerFactor_VINA_MEJORADA("15");
                traerFactor_FRANQUETTE("8");
                traerFactor_DESPELONADO("16");
                traerFactor_QUEBRADO("14");
            }
          
        }
     
        private void FormFactor_Load(object sender, EventArgs e)
        {
          
                traerFactor("2");
                traerFactor_Chandler("5");
                traerFactor_HOWARD("9");
                traerFactor_TULARE("10");
                traerFactor_SUNDLAND("6");
                traerFactor_HARTLEY("4");
                traerFactor_SEMILLA("1");
                traerFactor_VINA("3");
                traerFactor_VINA_MEJORADA("15");
                traerFactor_FRANQUETTE("8");
                traerFactor_DESPELONADO("16");
                traerFactor_QUEBRADO("14");
                dataGridView1.Columns[1].ReadOnly = true;
                dataGridView1.Columns[2].ReadOnly = true;
                dataGridView1.Columns["Id"].Visible = false;

                dataGridView2.Columns[1].ReadOnly = true;
                dataGridView2.Columns[2].ReadOnly = true;
                dataGridView2.Columns["Id"].Visible = false;

                dataGridView3.Columns[1].ReadOnly = true;
                dataGridView3.Columns[2].ReadOnly = true;
                dataGridView3.Columns["Id"].Visible = false;

                dataGridView4.Columns[1].ReadOnly = true;
                dataGridView4.Columns[2].ReadOnly = true;
                dataGridView4.Columns["Id"].Visible = false;

                dataGridView5.Columns[1].ReadOnly = true;
                dataGridView5.Columns[2].ReadOnly = true;
                dataGridView5.Columns["Id"].Visible = false;

                dataGridView6.Columns[1].ReadOnly = true;
                dataGridView6.Columns[2].ReadOnly = true;
                dataGridView6.Columns["Id"].Visible = false;

                dataGridView7.Columns[1].ReadOnly = true;
                dataGridView7.Columns[2].ReadOnly = true;
                dataGridView7.Columns["Id"].Visible = false;

                dataGridView9.Columns[1].ReadOnly = true;
                dataGridView9.Columns[2].ReadOnly = true;
                dataGridView9.Columns["Id"].Visible = false;

                dataGridView10.Columns[1].ReadOnly = true;
                dataGridView10.Columns[2].ReadOnly = true;
                dataGridView10.Columns["Id"].Visible = false;

                dataGridView11.Columns[1].ReadOnly = true;
                dataGridView11.Columns[2].ReadOnly = true;
                dataGridView11.Columns["Id"].Visible = false;

                dataGridView12.Columns[1].ReadOnly = true;
                dataGridView12.Columns[2].ReadOnly = true;
                dataGridView12.Columns["Id"].Visible = false;
            
            
      

            CmbProductor();
            Txt_Codigo.Text = "01";

        }

        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
          
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

            DialogResult dialogResult = MessageBox.Show("Seguro que desea modificar este dato ??", "Anakena", MessageBoxButtons.YesNo,MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {

                if((Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["% Estimacion"].Value.ToString())+Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["% Real"].Value.ToString())) <= 100)
                { 
                ModificarFactor(dataGridView1.Rows[e.RowIndex].Cells["ID"].Value.ToString(),Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["% Estimacion"].Value.ToString()),Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["% Real"].Value.ToString()));
                MessageBox.Show("Dato modificado correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);
               
                }
                else
                {
                    MessageBox.Show("La suma del porcentaje estimacion mas el real no pueden sumar mas de 100 % ??", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                traerFactor("2");
                traerFactor_Chandler("5");
                traerFactor_HOWARD("9");
                traerFactor_TULARE("10");
                traerFactor_SUNDLAND("6");
                traerFactor_HARTLEY("4");
                traerFactor_SEMILLA("1");
                traerFactor_VINA("3");
                traerFactor_VINA_MEJORADA("15");
                traerFactor_FRANQUETTE("8");
                traerFactor_DESPELONADO("16");
                traerFactor_QUEBRADO("14");
            }
            else if (dialogResult == DialogResult.No)
            {
              
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            tabControl1.Visible = false;
            traerFactor_filtro(Txt_Codigo.Text);
           
            dataGridView13.Visible = true;
            dataGridView13.Columns["Id"].Visible = false;


        }

        private void button2_Click(object sender, EventArgs e)
        {
            traerFactor("2");
            traerFactor_Chandler("5");
            traerFactor_HOWARD("9");
            traerFactor_TULARE("10");
            traerFactor_SUNDLAND("6");
            traerFactor_HARTLEY("4");
            traerFactor_SEMILLA("1");
            traerFactor_VINA("3");
            traerFactor_VINA_MEJORADA("15");
            traerFactor_FRANQUETTE("8");
            traerFactor_DESPELONADO("16");
            traerFactor_QUEBRADO("14");
            cmb_productor.SelectedIndex = 0;
            dataGridView13.Visible = false;
            tabControl1.Visible = true;
        }

        private void dataGridView1_CellEndEdit_1(object sender, DataGridViewCellEventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Seguro que desea modificar este dato ??", "Anakena", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {

                //if ((Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["% Estimacion"].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["% Real"].Value.ToString())) <= 100)
                if(dataGridView1.Columns[e.ColumnIndex].Name.ToString() == "% Estimacion" )
                {
                    double estimacion = Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["% Estimacion"].Value.ToString());
                    ModificarFactor(dataGridView1.Rows[e.RowIndex].Cells["ID"].Value.ToString(),estimacion,100-estimacion);
                    MessageBox.Show("Dato modificado correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                else
                 if (dataGridView1.Columns[e.ColumnIndex].Name.ToString() == "% Real")
                {
                    double Real = Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["% Real"].Value.ToString());
                    ModificarFactor(dataGridView1.Rows[e.RowIndex].Cells["ID"].Value.ToString(), 100- Real, Real);
                    MessageBox.Show("Dato modificado correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                traerFactor("2");
            }
           
        }

      

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            FormBusquedaProductor s = new FormBusquedaProductor();
            s.ShowDialog();
            Txt_Codigo.Text = "";
            cmb_productor.Text = s.nombre;
            Txt_Codigo.Text = s.codigo;
        }

        private void cmb_productor_SelectedIndexChanged(object sender, EventArgs e)
        {
            try {
                Txt_Codigo.Text = cmb_productor.SelectedValue.ToString();
            }
            catch { }
        }

        private void Txt_Codigo_KeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {
                if (e.KeyChar == Convert.ToChar(Keys.Enter))
                {
     

                    cmb_productor.SelectedIndex =((Convert.ToInt32(Txt_Codigo.Text)) - 1);
                   
                }
               
        
                 }
            catch { }
        }

        private void dataGridView2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Seguro que desea modificar este dato ??", "Anakena", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {

                //if ((Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["% Estimacion"].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["% Real"].Value.ToString())) <= 100)
                if (dataGridView2.Columns[e.ColumnIndex].Name.ToString() == "% Estimacion")
                {
                    double estimacion = Convert.ToDouble(dataGridView2.Rows[e.RowIndex].Cells["% Estimacion"].Value.ToString());
                    ModificarFactor(dataGridView2.Rows[e.RowIndex].Cells["ID"].Value.ToString(), estimacion, 100 - estimacion);
                    MessageBox.Show("Dato modificado correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                else
                 if (dataGridView2.Columns[e.ColumnIndex].Name.ToString() == "% Real")
                {
                    double Real = Convert.ToDouble(dataGridView2.Rows[e.RowIndex].Cells["% Real"].Value.ToString());
                    ModificarFactor(dataGridView2.Rows[e.RowIndex].Cells["ID"].Value.ToString(), 100 - Real, Real);
                    MessageBox.Show("Dato modificado correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                traerFactor("2");
                traerFactor("2");
                traerFactor_Chandler("5");
                traerFactor_HOWARD("9");
                traerFactor_TULARE("10");
                traerFactor_SUNDLAND("6");
                traerFactor_HARTLEY("4");
                traerFactor_SEMILLA("1");
                traerFactor_VINA("3");
                traerFactor_VINA_MEJORADA("15");
                traerFactor_FRANQUETTE("8");
                traerFactor_DESPELONADO("16");
                traerFactor_QUEBRADO("14");
            }
        }

        private void dataGridView3_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Seguro que desea modificar este dato ??", "Anakena", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {

                //if ((Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["% Estimacion"].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["% Real"].Value.ToString())) <= 100)
                if (dataGridView3.Columns[e.ColumnIndex].Name.ToString() == "% Estimacion")
                {
                    double estimacion = Convert.ToDouble(dataGridView3.Rows[e.RowIndex].Cells["% Estimacion"].Value.ToString());
                    ModificarFactor(dataGridView3.Rows[e.RowIndex].Cells["ID"].Value.ToString(), estimacion, 100 - estimacion);
                    MessageBox.Show("Dato modificado correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                else
                 if (dataGridView3.Columns[e.ColumnIndex].Name.ToString() == "% Real")
                {
                    double Real = Convert.ToDouble(dataGridView3.Rows[e.RowIndex].Cells["% Real"].Value.ToString());
                    ModificarFactor(dataGridView3.Rows[e.RowIndex].Cells["ID"].Value.ToString(), 100 - Real, Real);
                    MessageBox.Show("Dato modificado correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                traerFactor("2");
                traerFactor_Chandler("5");
                traerFactor_HOWARD("9");
                traerFactor_TULARE("10");
                traerFactor_SUNDLAND("6");
                traerFactor_HARTLEY("4");
                traerFactor_SEMILLA("1");
                traerFactor_VINA("3");
                traerFactor_VINA_MEJORADA("15");
                traerFactor_FRANQUETTE("8");
                traerFactor_DESPELONADO("16");
                traerFactor_QUEBRADO("14");
            }
        }

        private void dataGridView4_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Seguro que desea modificar este dato ??", "Anakena", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {

                //if ((Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["% Estimacion"].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["% Real"].Value.ToString())) <= 100)
                if (dataGridView4.Columns[e.ColumnIndex].Name.ToString() == "% Estimacion")
                {
                    double estimacion = Convert.ToDouble(dataGridView4.Rows[e.RowIndex].Cells["% Estimacion"].Value.ToString());
                    ModificarFactor(dataGridView4.Rows[e.RowIndex].Cells["ID"].Value.ToString(), estimacion, 100 - estimacion);
                    MessageBox.Show("Dato modificado correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                else
                 if (dataGridView4.Columns[e.ColumnIndex].Name.ToString() == "% Real")
                {
                    double Real = Convert.ToDouble(dataGridView4.Rows[e.RowIndex].Cells["% Real"].Value.ToString());
                    ModificarFactor(dataGridView4.Rows[e.RowIndex].Cells["ID"].Value.ToString(), 100 - Real, Real);
                    MessageBox.Show("Dato modificado correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                traerFactor("2");
                traerFactor_Chandler("5");
                traerFactor_HOWARD("9");
                traerFactor_TULARE("10");
                traerFactor_SUNDLAND("6");
                traerFactor_HARTLEY("4");
                traerFactor_SEMILLA("1");
                traerFactor_VINA("3");
                traerFactor_VINA_MEJORADA("15");
                traerFactor_FRANQUETTE("8");
                traerFactor_DESPELONADO("16");
                traerFactor_QUEBRADO("14");
            }
        }

   

        private void dataGridView6_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

            DialogResult dialogResult = MessageBox.Show("Seguro que desea modificar este dato ??", "Anakena", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {

              
                if (dataGridView6.Columns[e.ColumnIndex].Name.ToString() == "% Estimacion")
                {
                    double estimacion = Convert.ToDouble(dataGridView6.Rows[e.RowIndex].Cells["% Estimacion"].Value.ToString());
                    ModificarFactor(dataGridView6.Rows[e.RowIndex].Cells["ID"].Value.ToString(), estimacion, 100 - estimacion);
                    MessageBox.Show("Dato modificado correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                else
                 if (dataGridView6.Columns[e.ColumnIndex].Name.ToString() == "% Real")
                {
                    double Real = Convert.ToDouble(dataGridView6.Rows[e.RowIndex].Cells["% Real"].Value.ToString());
                    ModificarFactor(dataGridView6.Rows[e.RowIndex].Cells["ID"].Value.ToString(), 100 - Real, Real);
                    MessageBox.Show("Dato modificado correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                traerFactor("2");
                traerFactor_Chandler("5");
                traerFactor_HOWARD("9");
                traerFactor_TULARE("10");
                traerFactor_SUNDLAND("6");
                traerFactor_HARTLEY("4");
                traerFactor_SEMILLA("1");
                traerFactor_VINA("3");
                traerFactor_VINA_MEJORADA("15");
                traerFactor_FRANQUETTE("8");
                traerFactor_DESPELONADO("16");
                traerFactor_QUEBRADO("14");
            }
        }

        private void dataGridView7_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

            DialogResult dialogResult = MessageBox.Show("Seguro que desea modificar este dato ??", "Anakena", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {

                //if ((Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["% Estimacion"].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["% Real"].Value.ToString())) <= 100)
                if (dataGridView7.Columns[e.ColumnIndex].Name.ToString() == "% Estimacion")
                {
                    double estimacion = Convert.ToDouble(dataGridView7.Rows[e.RowIndex].Cells["% Estimacion"].Value.ToString());
                    ModificarFactor(dataGridView7.Rows[e.RowIndex].Cells["ID"].Value.ToString(), estimacion, 100 - estimacion);
                    MessageBox.Show("Dato modificado correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                else
                 if (dataGridView7.Columns[e.ColumnIndex].Name.ToString() == "% Real")
                {
                    double Real = Convert.ToDouble(dataGridView7.Rows[e.RowIndex].Cells["% Real"].Value.ToString());
                    ModificarFactor(dataGridView7.Rows[e.RowIndex].Cells["ID"].Value.ToString(), 100 - Real, Real);
                    MessageBox.Show("Dato modificado correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                traerFactor("2");
                traerFactor_Chandler("5");
                traerFactor_HOWARD("9");
                traerFactor_TULARE("10");
                traerFactor_SUNDLAND("6");
                traerFactor_HARTLEY("4");
                traerFactor_SEMILLA("1");
                traerFactor_VINA("3");
                traerFactor_VINA_MEJORADA("15");
                traerFactor_FRANQUETTE("8");
                traerFactor_DESPELONADO("16");
                traerFactor_QUEBRADO("14");
            }
        }

        private void dataGridView8_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

            DialogResult dialogResult = MessageBox.Show("Seguro que desea modificar este dato ??", "Anakena", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {

                //if ((Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["% Estimacion"].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["% Real"].Value.ToString())) <= 100)
                if (dataGridView8.Columns[e.ColumnIndex].Name.ToString() == "% Estimacion")
                {
                    double estimacion = Convert.ToDouble(dataGridView8.Rows[e.RowIndex].Cells["% Estimacion"].Value.ToString());
                    ModificarFactor(dataGridView8.Rows[e.RowIndex].Cells["ID"].Value.ToString(), estimacion, 100 - estimacion);
                    MessageBox.Show("Dato modificado correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                else
                 if (dataGridView8.Columns[e.ColumnIndex].Name.ToString() == "% Real")
                {
                    double Real = Convert.ToDouble(dataGridView8.Rows[e.RowIndex].Cells["% Real"].Value.ToString());
                    ModificarFactor(dataGridView8.Rows[e.RowIndex].Cells["ID"].Value.ToString(), 100 - Real, Real);
                    MessageBox.Show("Dato modificado correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                traerFactor("2");
                traerFactor_Chandler("5");
                traerFactor_HOWARD("9");
                traerFactor_TULARE("10");
                traerFactor_SUNDLAND("6");
                traerFactor_HARTLEY("4");
                traerFactor_SEMILLA("1");
                traerFactor_VINA("3");
                traerFactor_VINA_MEJORADA("15");
                traerFactor_FRANQUETTE("8");
                traerFactor_DESPELONADO("16");
                traerFactor_QUEBRADO("14");
            }
        }

        private void dataGridView9_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

            DialogResult dialogResult = MessageBox.Show("Seguro que desea modificar este dato ??", "Anakena", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {

                if (dataGridView9.Columns[e.ColumnIndex].Name.ToString() == "% Estimacion")
                {
                    double estimacion = Convert.ToDouble(dataGridView9.Rows[e.RowIndex].Cells["% Estimacion"].Value.ToString());
                    ModificarFactor(dataGridView9.Rows[e.RowIndex].Cells["ID"].Value.ToString(), estimacion, 100 - estimacion);
                    MessageBox.Show("Dato modificado correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                else
                 if (dataGridView9.Columns[e.ColumnIndex].Name.ToString() == "% Real")
                {
                    double Real = Convert.ToDouble(dataGridView9.Rows[e.RowIndex].Cells["% Real"].Value.ToString());
                    ModificarFactor(dataGridView9.Rows[e.RowIndex].Cells["ID"].Value.ToString(), 100 - Real, Real);
                    MessageBox.Show("Dato modificado correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                traerFactor("2");
                traerFactor_Chandler("5");
                traerFactor_HOWARD("9");
                traerFactor_TULARE("10");
                traerFactor_SUNDLAND("6");
                traerFactor_HARTLEY("4");
                traerFactor_SEMILLA("1");
                traerFactor_VINA("3");
                traerFactor_VINA_MEJORADA("15");
                traerFactor_FRANQUETTE("8");
                traerFactor_DESPELONADO("16");
                traerFactor_QUEBRADO("14");
            }
        }

        private void dataGridView10_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

            DialogResult dialogResult = MessageBox.Show("Seguro que desea modificar este dato ??", "Anakena", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {

                //if ((Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["% Estimacion"].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["% Real"].Value.ToString())) <= 100)
                if (dataGridView10.Columns[e.ColumnIndex].Name.ToString() == "% Estimacion")
                {
                    double estimacion = Convert.ToDouble(dataGridView10.Rows[e.RowIndex].Cells["% Estimacion"].Value.ToString());
                    ModificarFactor(dataGridView10.Rows[e.RowIndex].Cells["ID"].Value.ToString(), estimacion, 100 - estimacion);
                    MessageBox.Show("Dato modificado correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                else
                 if (dataGridView10.Columns[e.ColumnIndex].Name.ToString() == "% Real")
                {
                    double Real = Convert.ToDouble(dataGridView10.Rows[e.RowIndex].Cells["% Real"].Value.ToString());
                    ModificarFactor(dataGridView10.Rows[e.RowIndex].Cells["ID"].Value.ToString(), 100 - Real, Real);
                    MessageBox.Show("Dato modificado correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                traerFactor("2");
                traerFactor_Chandler("5");
                traerFactor_HOWARD("9");
                traerFactor_TULARE("10");
                traerFactor_SUNDLAND("6");
                traerFactor_HARTLEY("4");
                traerFactor_SEMILLA("1");
                traerFactor_VINA("3");
                traerFactor_VINA_MEJORADA("15");
                traerFactor_FRANQUETTE("8");
                traerFactor_DESPELONADO("16");
                traerFactor_QUEBRADO("14");
            }
        }

        private void dataGridView11_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

            DialogResult dialogResult = MessageBox.Show("Seguro que desea modificar este dato ??", "Anakena", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {

                //if ((Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["% Estimacion"].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["% Real"].Value.ToString())) <= 100)
                if (dataGridView11.Columns[e.ColumnIndex].Name.ToString() == "% Estimacion")
                {
                    double estimacion = Convert.ToDouble(dataGridView11.Rows[e.RowIndex].Cells["% Estimacion"].Value.ToString());
                    ModificarFactor(dataGridView11.Rows[e.RowIndex].Cells["ID"].Value.ToString(), estimacion, 100 - estimacion);
                    MessageBox.Show("Dato modificado correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                else
                 if (dataGridView11.Columns[e.ColumnIndex].Name.ToString() == "% Real")
                {
                    double Real = Convert.ToDouble(dataGridView11.Rows[e.RowIndex].Cells["% Real"].Value.ToString());
                    ModificarFactor(dataGridView11.Rows[e.RowIndex].Cells["ID"].Value.ToString(), 100 - Real, Real);
                    MessageBox.Show("Dato modificado correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                traerFactor("2");
                traerFactor_Chandler("5");
                traerFactor_HOWARD("9");
                traerFactor_TULARE("10");
                traerFactor_SUNDLAND("6");
                traerFactor_HARTLEY("4");
                traerFactor_SEMILLA("1");
                traerFactor_VINA("3");
                traerFactor_VINA_MEJORADA("15");
                traerFactor_FRANQUETTE("8");
                traerFactor_DESPELONADO("16");
                traerFactor_QUEBRADO("14");
            }
        }

        private void dataGridView12_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

            DialogResult dialogResult = MessageBox.Show("Seguro que desea modificar este dato ??", "Anakena", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {

                //if ((Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["% Estimacion"].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["% Real"].Value.ToString())) <= 100)
                if (dataGridView12.Columns[e.ColumnIndex].Name.ToString() == "% Estimacion")
                {
                    double estimacion = Convert.ToDouble(dataGridView12.Rows[e.RowIndex].Cells["% Estimacion"].Value.ToString());
                    ModificarFactor(dataGridView12.Rows[e.RowIndex].Cells["ID"].Value.ToString(), estimacion, 100 - estimacion);
                    MessageBox.Show("Dato modificado correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                else
                 if (dataGridView12.Columns[e.ColumnIndex].Name.ToString() == "% Real")
                {
                    double Real = Convert.ToDouble(dataGridView12.Rows[e.RowIndex].Cells["% Real"].Value.ToString());
                    ModificarFactor(dataGridView12.Rows[e.RowIndex].Cells["ID"].Value.ToString(), 100 - Real, Real);
                    MessageBox.Show("Dato modificado correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                traerFactor("2");
                traerFactor_Chandler("5");
                traerFactor_HOWARD("9");
                traerFactor_TULARE("10");
                traerFactor_SUNDLAND("6");
                traerFactor_HARTLEY("4");
                traerFactor_SEMILLA("1");
                traerFactor_VINA("3");
                traerFactor_VINA_MEJORADA("15");
                traerFactor_FRANQUETTE("8");
                traerFactor_DESPELONADO("16");
                traerFactor_QUEBRADO("14");
            }
        }

        private void dataGridView5_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

            DialogResult dialogResult = MessageBox.Show("Seguro que desea modificar este dato ??", "Anakena", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {

                //if ((Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["% Estimacion"].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["% Real"].Value.ToString())) <= 100)
                if (dataGridView5.Columns[e.ColumnIndex].Name.ToString() == "% Estimacion")
                {
                    double estimacion = Convert.ToDouble(dataGridView5.Rows[e.RowIndex].Cells["% Estimacion"].Value.ToString());
                    ModificarFactor(dataGridView5.Rows[e.RowIndex].Cells["ID"].Value.ToString(), estimacion, 100 - estimacion);
                    MessageBox.Show("Dato modificado correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                else
                 if (dataGridView5.Columns[e.ColumnIndex].Name.ToString() == "% Real")
                {
                    double Real = Convert.ToDouble(dataGridView5.Rows[e.RowIndex].Cells["% Real"].Value.ToString());
                    ModificarFactor(dataGridView5.Rows[e.RowIndex].Cells["ID"].Value.ToString(), 100 - Real, Real);
                    MessageBox.Show("Dato modificado correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                traerFactor("2");
                traerFactor_Chandler("5");
                traerFactor_HOWARD("9");
                traerFactor_TULARE("10");
                traerFactor_SUNDLAND("6");
                traerFactor_HARTLEY("4");
                traerFactor_SEMILLA("1");
                traerFactor_VINA("3");
                traerFactor_VINA_MEJORADA("15");
                traerFactor_FRANQUETTE("8");
                traerFactor_DESPELONADO("16");
                traerFactor_QUEBRADO("14");
            }
        }
    }

}
