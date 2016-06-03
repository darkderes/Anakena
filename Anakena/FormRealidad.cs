using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.Collections;

namespace Anakena
{
    public partial class FormRealidad : Form
    {
        conexion ex = new conexion();
        public FormRealidad()
        {
            InitializeComponent();
        }

        private void FormRealidad_Load(object sender, EventArgs e)
        {
            
            CmbVariedad();
            radioButton1.Checked = true;
            dataGridView4.Visible = false;
            dataGridView5.Visible = false;
            dataGridView6.Visible = false;
            lblActuallizacion.Text = lblActuallizacion.Text + Traer_Fecha();
    

        }
        public string Traer_Fecha()
        {
            String id = "";
            SqlCommand cmd = new SqlCommand("spTraer_Fecha_real", ex.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            ex.Abrir();
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
            ex.Cerrar();
            return id;
        }
        public void traer_estimacion()
        {
            SqlCommand cmd = new SqlCommand("spTraer_Estimacion2", ex.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.VarChar, 25);
            cmd.Parameters["@variedad"].Value = cmb_variedad.SelectedValue;
            ex.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView6.DataSource = myds.Tables[0];
            ex.Cerrar();

        }
        public void RealPORC()
        {
            SqlCommand cmd = new SqlCommand("spTraer_Estimacion2%", ex.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Int);
            cmd.Parameters["@variedad"].Value = cmb_variedad.SelectedValue.ToString();
            ex.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView6.DataSource = myds.Tables[0];
            ex.Cerrar();

        }
        public void Estimacion_CalibreKG()
        {
            SqlCommand cmd = new SqlCommand("spRealCalibreKilos", ex.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Int);
            cmd.Parameters["@variedad"].Value = cmb_variedad.SelectedValue.ToString();
            ex.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView5.DataSource = myds.Tables[0];
            ex.Cerrar();
        }
        public void Estimacion_CalibrePORC()
        {
            SqlCommand cmd = new SqlCommand("spRealCalibre%", ex.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Int);
            cmd.Parameters["@variedad"].Value = cmb_variedad.SelectedValue.ToString();
            ex.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView5.DataSource = myds.Tables[0];
            ex.Cerrar();

        }
        public void Estimacion_CategoriasKG()
        {
            SqlCommand cmd = new SqlCommand("spRealCategoriasKilos", ex.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Int);
            cmd.Parameters["@variedad"].Value = cmb_variedad.SelectedValue.ToString();
            ex.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView4.DataSource = myds.Tables[0];
            ex.Cerrar();

        }
        public void Estimacion_CategoriaPORC()
        {
            SqlCommand cmd = new SqlCommand("spRealCategorias%", ex.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Int);
            cmd.Parameters["@variedad"].Value = cmb_variedad.SelectedValue.ToString();
            ex.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView4.DataSource = myds.Tables[0];
            ex.Cerrar();

        }
        public void CmbVariedad()
        {
            try
            {
                ex.Abrir();
                SqlCommand command = new SqlCommand("spTraerVariedad", ex.getConexion());
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
                ex.Cerrar();
            }
        }
        private void BtnFiltro_Click(object sender, EventArgs e)
        {
            if(radioButton1.Checked == true)
            {
                traer_estimacion();  
                Estimacion_CalibreKG();       
                Estimacion_CategoriasKG();
            }
            else
            {
                RealPORC();
                Estimacion_CalibrePORC();
                Estimacion_CategoriaPORC();
            }
            dataGridView6.Visible = true;
            dataGridView5.Visible = true;
            dataGridView4.Visible = true;
            pictureBox2.Visible = true;
            pictureBox3.Visible = false;
            pictureBox1.Visible = false;
            dataGridView6.Columns[0].Visible = false;
            dataGridView6.Columns[1].Visible = false;
            dataGridView5.Columns[0].Visible = false;
            dataGridView4.Columns[0].Visible = false;

        }
        public void Update_Real()
        {
            try
            {
                // linea de comando de sql
                SqlCommand cmd = new SqlCommand("spExtraer_Pall", ex.getConexion());

                // adhrsion de parametros 
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@msg", SqlDbType.VarChar, 100);
                cmd.Parameters["@msg"].Value = 1;
                //abre coneccion a base de datos
                ex.Abrir();
                cmd.ExecuteNonQuery();
                string sqlMensaje = cmd.Parameters["@msg"].Value.ToString();
                //cierre de coneccion a base de datos
                ex.Cerrar();
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
        public void exporta_a_excel_categoria()
        {
            ApplicationClass class2 = new ApplicationClass();
            class2.Application.Workbooks.Add(true);
            int num = 0;
            int numC = 0;
            pBar1.Minimum = 1;
            pBar1.Maximum = dataGridView5.RowCount * dataGridView5.ColumnCount;
            pBar1.Value = 1;
            pBar1.Step = 1;

            foreach (DataGridViewColumn column in this.dataGridView5.Columns)
            {
                class2.Cells[1, 1] = cmb_variedad.Text.Trim();
                num++;
                class2.Cells[1, num] = column.Name;
                class2.get_Range("A1", "GF1").Interior.ColorIndex = 9;
                class2.get_Range("A1", "GF1").Font.ColorIndex = 2;
            }

            foreach (DataGridViewColumn column in this.dataGridView6.Columns)
            {
                class2.Cells[10, 1] = cmb_variedad.Text.Trim();
                numC++;
                class2.Cells[10, numC] = column.Name;
                class2.get_Range("A10", "GF10").Interior.ColorIndex = 9;
                class2.get_Range("A10", "GF10").Font.ColorIndex = 2;
            }
            int num22 = 0;
            foreach (DataGridViewRow row in (IEnumerable)this.dataGridView5.Rows)
            {
                num22++;
                num = 0;
                foreach (DataGridViewColumn column2 in this.dataGridView6.Columns)
                {
                    num++;
                    class2.Cells[num22 + 1, num] = row.Cells[column2.Name].Value;
                    pBar1.PerformStep();
                    label2.Text = "Proceso :" + pBar1.Value.ToString() + "/" + (dataGridView5.RowCount * dataGridView5.ColumnCount).ToString();
                }
            }
            int num2 =9;
            foreach (DataGridViewRow row in (IEnumerable)this.dataGridView6.Rows)
            {
                num2++;
                num= 0;
                foreach (DataGridViewColumn column2 in this.dataGridView6.Columns)
                {
                    num++;
                    class2.Cells[num2+1, num] = row.Cells[column2.Name].Value;
                    pBar1.PerformStep();

                }
            }
            class2.Visible = true;
            ((Worksheet)class2.ActiveSheet).Name = cmb_variedad.Text.Trim();
            ((Worksheet)class2.ActiveSheet).Activate();

        }
        //private void Btn_Excel_Click(object sender, EventArgs e)
        //{
        //    if (dataGridView2.Visible == true)
        //    {
        //        dataGridView2.Visible = false;
        //    exporta_a_excel();
        //    }
           
        //    else
        //    {
        //        pictureBox1.Visible = true;
        //        dataGridView5.Visible = false;
        //        dataGridView6.Visible = false;

        //            exporta_a_excel_categoria();
        //    }
        //    pBar1.Visible = false; Btn_Excel.Visible = false;
        //}

       private void dataGridView4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dataGridView5_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            try
            {
                if (dataGridView5.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString() == "")
                {
                    dataGridView5.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = 0;
                }
            }
            catch
            { }
        }

        private void dataGridView6_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            try
            {
                if (dataGridView6.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString() == "")
                {
                    dataGridView6.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = 0;
                }
            }
            catch
            { }
          
        }
      

        private void dataGridView4_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            try
            {
                if (dataGridView4.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString() == "")
                {
                    dataGridView4.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = 0;
                }
            }
            catch
            { }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Update_Real();
                Traer_Fecha();
                MessageBox.Show("Datos actualizados correctamente","Anakena",MessageBoxButtons.OK,MessageBoxIcon.Information);
                
            }
            catch {
                MessageBox.Show("Ocurrio un problema al tratar de actualizar datos", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView6_Scroll(object sender, ScrollEventArgs e)
        {
            dataGridView6.Columns[2].Frozen = true;
            dataGridView4.HorizontalScrollingOffset = e.NewValue;
            dataGridView5.HorizontalScrollingOffset = e.NewValue;
        }

        private void dataGridView5_Scroll(object sender, ScrollEventArgs e)
        {
            dataGridView5.Columns[1].Frozen = true;
            dataGridView4.HorizontalScrollingOffset = e.NewValue;
            dataGridView6.HorizontalScrollingOffset = e.NewValue;
        }

        private void dataGridView4_Scroll(object sender, ScrollEventArgs e)
        {
            dataGridView4.Columns[1].Frozen = true;
            dataGridView5.HorizontalScrollingOffset = e.NewValue;
            dataGridView6.HorizontalScrollingOffset = e.NewValue;
        }
    }
    
}
