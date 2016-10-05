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
    public partial class FormRendimiento_Descarte : Form
    {
        conexion cn = new conexion();
        public FormRendimiento_Descarte()
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
                //importarDatosSERR(TxtFile.Text);
                importarDatosCHANDLER(TxtFile.Text);
                //importarDatosHOWARD(TxtFile.Text);
                //importarDatosFRANQUETTE(TxtFile.Text);
                //importarDatosHARTLEY(TxtFile.Text);
                //importarDatosSEMILLA(TxtFile.Text);
                //importarDatosSUNDLAND(TxtFile.Text);
                //importarDatosTULARE(TxtFile.Text);
                //importarDatosVINA(TxtFile.Text);
                //importarDatosVINA_MEJ(TxtFile.Text);

                if (dataGridView2.Rows.Count > 0)
                {
                    btn_ingresar.Visible = true;
                    Btn_Examinar.Visible = false;
                }
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
            catch(Exception ex)
            {
                MessageBox.Show("Error al tratar cargar Hoja 'CHANDLER' verifique que existe en el archivo : "+ex, "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }         
        }
        public void Agregar_Rendimiento_Descarte(string Producto, int variedad, float porcentajes)
        {
            try
            {
                // linea de comando de sql
                SqlCommand cmd = new SqlCommand("spAgregar_Rendimiento_Descarte", cn.getConexion());
                // adhesion de parametros 
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@Producto", SqlDbType.Text);
                cmd.Parameters.Add("@Cod_Variedad", SqlDbType.Char, 2);
                cmd.Parameters.Add("@porcentajes", SqlDbType.Float);
                cmd.Parameters.Add("@msg", SqlDbType.VarChar, 100);
                // ingreso de parametros
                cmd.Parameters["@Producto"].Value = Producto;
                cmd.Parameters["@Cod_Variedad"].Value = variedad;
                cmd.Parameters["@porcentajes"].Value = porcentajes;
                cmd.Parameters["@msg"].Value = 1;
                //abre coneccion a base de datos
                cn.Abrir();
                cmd.ExecuteNonQuery();
                //cierre de coneccion a base de datos
                cn.Cerrar();
             
            }
            catch (Exception ex)
            {
                //mensaje de error SQL
                MessageBox.Show("Error SQL " + ex, "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        public void Eliminar_Rendimiento_Descarte()
        {
            try
            {
                // linea de comando de sql
                SqlCommand cmd = new SqlCommand("spDelete_Rendimiento_Descarte", cn.getConexion());
                // adhesion de parametros 
                cmd.CommandType = CommandType.StoredProcedure;
           
                cmd.Parameters.Add("@msg", SqlDbType.VarChar, 100);
                // ingreso de parametros
         
                cmd.Parameters["@msg"].Value = 1;
                //abre coneccion a base de datos
                cn.Abrir();
                cmd.ExecuteNonQuery();
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
            Eliminar_Rendimiento_Descarte();
            for (int x = 0; x < dataGridView2.RowCount; x++)
            {
                Agregar_Rendimiento_Descarte((dataGridView2.Rows[x].Cells[0].Value.ToString()), 5, Convert.ToSingle(dataGridView2.Rows[x].Cells[1].Value.ToString()));
            }
            MessageBox.Show("Rendimiento ingresado correctamente", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
