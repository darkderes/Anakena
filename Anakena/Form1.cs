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

   
    public partial class Form1 : Form
    {
        conexion cn = new conexion();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: esta línea de código carga datos en la tabla 'anakenaDataSet.Especie' Puede moverla o quitarla según sea necesario.
    
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

        private void Form1_Load_1(object sender, EventArgs e)
        {

        }

        public void ingresar_estimacion(string cod, float serr, float chandler, float  howard, float tulare, float sundland, float hartley, float semil, float vina, float vina_mej, float franquette, float desc_desp, float desc_queb)
        {
         
                try
                {
                    // linea de comando de sql
                    SqlCommand cmd = new SqlCommand("spAgregaEstimacion", cn.getConexion());

                    // adhrsion de parametros 
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add("@cod", SqlDbType.Char, 10);
                    cmd.Parameters.Add("@Serr", SqlDbType.Float);
                    cmd.Parameters.Add("@Chandler", SqlDbType.Decimal,18);
                    cmd.Parameters.Add("@Howard", SqlDbType.Decimal, 18);
                    cmd.Parameters.Add("@Tulare", SqlDbType.Decimal, 18);
                    cmd.Parameters.Add("@Sundland", SqlDbType.Decimal, 18);
                    cmd.Parameters.Add("@Hartley", SqlDbType.Decimal, 18);
                    cmd.Parameters.Add("@Semil", SqlDbType.Decimal, 18);
                    cmd.Parameters.Add("@Vina", SqlDbType.Decimal, 18);
                    cmd.Parameters.Add("@Vina_mej", SqlDbType.Decimal, 18);
                    cmd.Parameters.Add("@Franquette", SqlDbType.Decimal, 18);
                    cmd.Parameters.Add("@Descte_desp", SqlDbType.Decimal, 18);
                    cmd.Parameters.Add("@Descte_queb", SqlDbType.Decimal, 18);

                    cmd.Parameters.Add("@msg", SqlDbType.VarChar, 100);

                    // ingreso de parametros
                    cmd.Parameters["@cod"].Value = cod;
                    cmd.Parameters["@Serr"].Value = serr;
                    cmd.Parameters["@Chandler"].Value = chandler;
                    cmd.Parameters["@Howard"].Value = howard;
                    cmd.Parameters["@Tulare"].Value = tulare;
                    cmd.Parameters["@Sundland"].Value = sundland;
                    cmd.Parameters["@Hartley"].Value = hartley;
                    cmd.Parameters["@Semil"].Value = semil;
                    cmd.Parameters["@Vina"].Value = vina;
                    cmd.Parameters["@Vina_mej"].Value = vina_mej;
                    cmd.Parameters["@Franquette"].Value = franquette;
                    cmd.Parameters["@Descte_desp"].Value = desc_desp;
                    cmd.Parameters["@Descte_queb"].Value = desc_queb;

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
              //  formato();
            }
        }
       /* public void formato2()
        {
               for(int i = 0; i< dataGridView1.RowCount;i++)
            {
                if(dataGridView1.Rows[i].Cells["SERR"].Value.ToString() != "0")
                {
                    dataGridView1.Rows[i].Cells["SERR"].Style.Format = "##,##.00";
                }

            }
        }*/
        public void formato()
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells["SERR"].Value.ToString() != "0")
                {
                    row.Cells["SERR"].Style.Format = "##,##.00";
                }
                if (row.Cells["CHANDLER"].Value.ToString() != "0")
                {
                    row.Cells["CHANDLER"].Style.Format = "##,##.00";
                }
                if (row.Cells["HOWARD"].Value.ToString() != "0")
                {
                    row.Cells["HOWARD"].Style.Format = "##,##.00";
                }
                if (row.Cells["TULARE"].Value.ToString() != "0")
                {
                    row.Cells["TULARE"].Style.Format = "##,##.00";
                }
                if (row.Cells["SUNLAND"].Value.ToString() != "0")
                {
                    row.Cells["SUNLAND"].Style.Format = "##,##.00";
                }
                if (row.Cells["HARTLEY"].Value.ToString() != "0")
                {
                    row.Cells["HARTLEY"].Style.Format = "##,##.00";
                }
                if (row.Cells["SEMILLA"].Value.ToString() != "0")
                {
                    row.Cells["SEMILLA"].Style.Format = "##,##.00";
                }
                if (row.Cells[9].Value.ToString() != "0")
                {
                    row.Cells[9].Style.Format = "##,##.00";
                }
                if (row.Cells["VINA MEJ"].Value.ToString() != "0")
                {
                    row.Cells["VINA MEJ"].Style.Format = "##,##.00";
                }
                if (row.Cells["FRANQUETTE"].Value.ToString() != "0")
                {
                    row.Cells["FRANQUETTE"].Style.Format = "##,##.00";
                }
                if (row.Cells["DESCTE DESP"].Value.ToString() != "0")
                {
                    row.Cells["DESCTE DESP"].Style.Format = "##,##.00";
                }
                if (row.Cells["DESCTE QUEB"].Value.ToString() != "0")
                {
                    row.Cells["DESCTE QUEB"].Style.Format = "##,##.00";
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
            eliminar_estimacion();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                string cod = row.Cells["cod"].Value.ToString();
                float serr = (float) Convert.ToSingle(row.Cells["SERR"].Value.ToString());
                float chandler = Convert.ToSingle(row.Cells["CHANDLER"].Value.ToString());
                float howard = Convert.ToSingle(row.Cells["HOWARD"].Value.ToString());
                float tulare = Convert.ToSingle(row.Cells["TULARE"].Value.ToString());
                float sundland = Convert.ToSingle(row.Cells["SUNLAND"].Value.ToString());
                float hartley = Convert.ToSingle(row.Cells["HARTLEY"].Value.ToString());
                float semil = Convert.ToSingle(row.Cells["SEMILLA"].Value.ToString());
                float vina = Convert.ToSingle(row.Cells[9].Value.ToString());
                float vina_mej = Convert.ToSingle(row.Cells["VINA MEJ"].Value.ToString());
                float franquette = Convert.ToSingle(row.Cells["FRANQUETTE"].Value.ToString());
                float desc_desp = Convert.ToSingle(row.Cells["DESCTE DESP"].Value.ToString());
                float desc_queb = Convert.ToSingle(row.Cells["DESCTE QUEB"].Value.ToString());
            

                ingresar_estimacion(cod,serr,chandler,howard,tulare,sundland,hartley,semil,vina,vina_mej,franquette,desc_desp,desc_queb);
            }
            MessageBox.Show("Estimacion ingresada correctamente","Anakena",MessageBoxButtons.OK,MessageBoxIcon.Information);
            this.Close();
        }

        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
          
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
         
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
