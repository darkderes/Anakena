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

   
    public partial class FormTemporada_Estimacion : Form
    {
        conexion cn = new conexion();
        public FormTemporada_Estimacion()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
    
        }

        public void importarDatos(string file)
        {
            try {
                OleDbCommand command = new OleDbCommand();
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + file + "';Extended Properties=Excel 8.0;";
                command = new OleDbCommand("select * from [Hoja1$]", conn);
                DataSet myds = new DataSet();
                OleDbDataAdapter adap = new OleDbDataAdapter(command);
                adap.Fill(myds);
                dataGridView1.DataSource = myds.Tables[0];
            }
            catch
            {
                MessageBox.Show("Asegurese de que eligio el archivo correcto y/o existe nombre de hoja Hoja1","Anakena",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }  
         
        }

        private void Form1_Load_1(object sender, EventArgs e)
        {

        }

        public void ingresar_estimacion(string cod, double monto, string  variedad)
        {
            

            try
                {
                    // linea de comando de sql
                    SqlCommand cmd = new SqlCommand("spAgregaEstimacion", cn.getConexion());

                    // adhrsion de parametros 
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add("@cod", SqlDbType.Char, 10);
                    cmd.Parameters.Add("@monto", SqlDbType.Float);
                    cmd.Parameters.Add("@variedad", SqlDbType.VarChar,20);
          

                    cmd.Parameters.Add("@msg", SqlDbType.VarChar, 100);

                    // ingreso de parametros
                    cmd.Parameters["@cod"].Value = cod;
                    cmd.Parameters["@monto"].Value = monto;
                    cmd.Parameters["@variedad"].Value = variedad;
                

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
          
        private void button1_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog newdialog = new OpenFileDialog();
            newdialog.Filter = "Excel files(*.xlsx)|*.xls";
            if (newdialog.ShowDialog() == DialogResult.OK)
            {
                TxtFile.Text = newdialog.FileName;
                importarDatos(TxtFile.Text);
                if(dataGridView1.Rows.Count > 0)
                {
                    btn_ingresar.Visible = true;
                    Btn_Examinar.Visible = false;
                }
        
            }
        }
   

        private void eliminar_estimacion()
        {
            try
            {
                // linea de comando de sql
                SqlCommand cmd = new SqlCommand("spDeleteEstimacion", cn.getConexion());

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
        private void btn_ingresar_Click(object sender, EventArgs e)
        {
           
            pBar1.Minimum = 1;
            pBar1.Maximum = (dataGridView1.RowCount * dataGridView1.ColumnCount);
            pBar1.Value = 1;
            pBar1.Step = 1;
            double total = 0;
            eliminar_estimacion();  
            foreach (DataGridViewColumn col in dataGridView1.Columns)
            {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                    pBar1.Visible = true;
                    pBar1.PerformStep();
                    if (col.Index > 1)
                    {
                       
                       string cod = row.Cells["cod"].Value.ToString();
                       string aux = row.Cells[col.Name.ToString()].Value.ToString();
                        if (String.IsNullOrEmpty(aux))
                        {
                            aux = "0";
                        }
                       double monto = Convert.ToDouble(aux);
                       total = total + monto;
                       string variedad = col.Name.ToString();
                        ingresar_estimacion(cod, monto, variedad);
                       
                    }


                }
            }
            pBar1.Visible = false;
            //lbl_cargando.Visible = false;
           MessageBox.Show("Estimacion ingresada correctamente con un total de "+total.ToString("##,##.00")+" Kg","Anakena",MessageBoxButtons.OK,MessageBoxIcon.Information);
            this.Close();
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.Value.ToString() != "0")
            {
                e.CellStyle.Format = "##,##.00";
            }
        }
    }
}
