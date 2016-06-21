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
using Microsoft.Office.Interop.Excel;
using System.Collections;
using System.IO;
using System.Threading;

namespace Anakena
{
    public partial class FormEstimacion : Form
    {
        conexion cn = new conexion();
        public FormEstimacion()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            CmbVariedad();
            radioButton1.Checked = true;
          


        }
        public void Estimacion_CalibreKG()
        {
            System.Windows.Forms.Application.DoEvents();
            SqlCommand cmd = new SqlCommand("spEstimacionCalibreKilos", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Int);
            cmd.Parameters["@variedad"].Value = cmb_variedad.SelectedValue.ToString();
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView4.DataSource = myds.Tables[0];
            cn.Cerrar();

        }
        public void Estimacion_CalibrePORC()
        {
          
            SqlCommand cmd = new SqlCommand("spEstimacionCalibre%", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Int);
            cmd.Parameters["@variedad"].Value = cmb_variedad.SelectedValue.ToString();
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView4.DataSource = myds.Tables[0];
            cn.Cerrar();

        }
        public void EstimacionKG()
        {
           
            SqlCommand cmd = new SqlCommand("spTraer_Estimacion", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Int);
            cmd.Parameters["@variedad"].Value = cmb_variedad.SelectedValue.ToString();
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView2.DataSource = myds.Tables[0];
            cn.Cerrar();

        }
        public void EstimacionPORC()
        {
            SqlCommand cmd = new SqlCommand("spTraer_Estimacion%", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Int);
            cmd.Parameters["@variedad"].Value = cmb_variedad.SelectedValue.ToString();
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView2.DataSource = myds.Tables[0];
            cn.Cerrar();

        }
        public void Estimacion_CategoriaKG()
        {
            SqlCommand cmd = new SqlCommand("spEstimacionCategoriasKilos", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Int);
            cmd.Parameters["@variedad"].Value = cmb_variedad.SelectedValue.ToString();
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView5.DataSource = myds.Tables[0];
            cn.Cerrar();

        }
        public void Estimacion_CategoriaPORC()
        {
            SqlCommand cmd = new SqlCommand("spEstimacionCategorias%", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Int);
            cmd.Parameters["@variedad"].Value = cmb_variedad.SelectedValue.ToString();
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView5.DataSource = myds.Tables[0];
            cn.Cerrar();

        }

        public void CmbVariedad()
        {
            try
            {
                cn.Abrir();
                SqlCommand command = new SqlCommand("spTraerVariedad", cn.getConexion());
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
                MessageBox.Show("Ocurrio un problema al cargar variedades","Anakena",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
            finally
            {
                cn.Cerrar();
            }
        }
  

        private void button1_Click(object sender, EventArgs e)
        {
            label2.Visible = true;
            pBar1.Visible = true;
            exporta_a_excel();
            label2.Visible = false;
            pBar1.Visible = false;
        }
        private void BtnFiltro_Click(object sender, EventArgs e)
        {

        
            pictureBox4.Image = Image.FromFile("C:/spin.gif");
            pictureBox4.Visible = true;
         
            

            dataGridView2.DataSource = null;
            dataGridView2.Rows.Clear();
            dataGridView4.DataSource = null;
            dataGridView4.Rows.Clear();
            dataGridView5.DataSource = null;
            dataGridView5.Rows.Clear();

            try
            {
                if ((cmb_variedad.SelectedValue.ToString() != "0") && (radioButton1.Checked == true))
                {
                  
                    Cursor = Cursors.WaitCursor;
                    System.Windows.Forms.Application.DoEvents();
                    Estimacion_CalibreKG();
                    System.Windows.Forms.Application.DoEvents();
                    Estimacion_CategoriaKG();
                    System.Windows.Forms.Application.DoEvents();
                    EstimacionKG();
                    System.Windows.Forms.Application.DoEvents();
                    dataGridView2.Columns[0].Visible = false;
                    dataGridView2.Columns[1].Visible = false;
                    dataGridView4.Columns[0].Visible = false;
                    dataGridView5.Columns[0].Visible = false;
                    dataGridView5.Columns[1].Width = 100;
                    Btn_Excel.Visible = true;
                    dataGridView2.Visible = true;
                    dataGridView4.Visible = true;
                    dataGridView5.Visible = true;
                    pictureBox1.Visible = true;
                    pictureBox2.Visible = false;
                    pictureBox3.Visible = false;
                    pBar1.Visible = false;
                    label2.Visible = false;
                    Cursor = Cursors.Default;
                }
                else
                if ((cmb_variedad.SelectedValue.ToString() != "0") && (radioButton2.Checked == true))
                {
                    Estimacion_CalibrePORC();
                    Estimacion_CategoriaPORC();
                    EstimacionPORC();
                    dataGridView4.Columns[0].Visible = false;
                    dataGridView5.Columns[0].Visible = false;
                    dataGridView2.Columns[0].Visible = false;
                    dataGridView2.Columns[1].Visible = false;
                    dataGridView4.Columns[0].Visible = false;
                    dataGridView5.Columns[0].Visible = false;
                    dataGridView5.Columns[1].Width = 100;
                    Btn_Excel.Visible = true;


                }
            }
            catch (Exception x) {
                MessageBox.Show("se produjo un error al tratar de cargar datos  "+ x.ToString());
            }
            pictureBox4.Visible = false;
            System.Windows.Forms.Application.DoEvents();

        }
        public void exporta_a_excel()
        {
            ApplicationClass class2 = new ApplicationClass();
            class2.Application.Workbooks.Add(true);
            int num = 0;
            int numC = 0;
            int numCC = 0;
            pBar1.Minimum = 1;
            pBar1.Maximum = dataGridView2.RowCount * dataGridView2.ColumnCount ;
            pBar1.Value = 1;
            pBar1.Step = 1;
       
            foreach (DataGridViewColumn column in this.dataGridView2.Columns)
            {
              class2.Cells[1, 1] = cmb_variedad.Text.Trim();
              num++;
              class2.Cells[1, num] = column.Name;
              class2.get_Range("A1", "GC1").Interior.ColorIndex = 9;
              class2.get_Range("A1", "GC1").Font.ColorIndex = 2;   
               
            }

            foreach (DataGridViewColumn column in this.dataGridView4.Columns)
            {
                class2.Cells[40, 1] = cmb_variedad.Text.Trim();
                numC++;
                class2.Cells[40, numC] = column.Name;
                class2.get_Range("A40", "GB40").Interior.ColorIndex = 9;
                class2.get_Range("A40", "GB40").Font.ColorIndex = 2;
            }
            foreach (DataGridViewColumn column in this.dataGridView5.Columns)
            {
                class2.Cells[54, 1] = cmb_variedad.Text.Trim();
                numCC++;
                class2.Cells[54, numCC] = column.Name;
                class2.get_Range("A54", "GB54").Interior.ColorIndex = 9;
                class2.get_Range("A54", "GB54").Font.ColorIndex = 2;
            }
            int num2 = 0;
            foreach (DataGridViewRow row in (IEnumerable)this.dataGridView2.Rows)
            {
                num2++;
                num = 0;
                foreach (DataGridViewColumn column2 in this.dataGridView2.Columns)
                {
                    num++;
                    class2.Cells[num2 + 1, num] = row.Cells[column2.Name].Value;    
                    pBar1.PerformStep();
                    label2.Text = "Proceso :" +pBar1.Value.ToString() +"/"+ (dataGridView2.RowCount*dataGridView2.ColumnCount).ToString();
                }
            }
            int num22 = 39;
            foreach (DataGridViewRow row in (IEnumerable)this.dataGridView4.Rows)
            {
                num22++;
                num = 0;
                foreach (DataGridViewColumn column2 in this.dataGridView4.Columns)
                {
                    num++;
                    class2.Cells[num22 + 1, num] = row.Cells[column2.Name].Value;
                    pBar1.PerformStep();

                }
            }
            int num222 = 53;
            foreach (DataGridViewRow row in (IEnumerable)this.dataGridView5.Rows)
            {
                num222++;
                num = 0;
                foreach (DataGridViewColumn column2 in this.dataGridView5.Columns)
                {
                    num++;
                    class2.Cells[num222 + 1, num] = row.Cells[column2.Name].Value;
                    pBar1.PerformStep();

                }
            }
            class2.Visible = true;
            ((Worksheet)class2.ActiveSheet).Name = cmb_variedad.Text.Trim();
            ((Worksheet)class2.ActiveSheet).Activate();
             
        }

        private void dataGridView4_Scroll(object sender, ScrollEventArgs e)
        {
            dataGridView5.HorizontalScrollingOffset = e.NewValue;
            dataGridView2.HorizontalScrollingOffset = e.NewValue;
            dataGridView4.Columns[1].Frozen = true;
        }

        private void dataGridView5_Scroll(object sender, ScrollEventArgs e)
        {
            dataGridView4.HorizontalScrollingOffset = e.NewValue;
            dataGridView2.HorizontalScrollingOffset = e.NewValue;
            dataGridView5.Columns[1].Frozen = true;
        }

        private void dataGridView2_Scroll(object sender, ScrollEventArgs e)
        {
            dataGridView2.Columns[2].Frozen = true;
            dataGridView4.HorizontalScrollingOffset = e.NewValue;
            dataGridView5.HorizontalScrollingOffset = e.NewValue;

        }

        private void dataGridView2_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
           
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }
    }
}
