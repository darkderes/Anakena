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
    public partial class FormRealidadVSEstimacion : Form
    {
        conexion ex = new conexion();
        conexionERP cn = new conexionERP();
        int cantidad = 0;
        public FormRealidadVSEstimacion()
        {
            InitializeComponent();
        }

        private void FormRealidadVSEstimacion_Load(object sender, EventArgs e)
        {
            CmbVariedad();
            CmbTipo.SelectedIndex = 0;
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.WorkerSupportsCancellation = true;
            //radioButton1.Checked = true;

        }
        //public void Estimacion_CalibrePORC()
        //{
        //    SqlCommand cmd = new SqlCommand("spEstimacionCalibre%", ex.getConexion());
        //    cmd.CommandType = CommandType.StoredProcedure;
        //    cmd.Parameters.Add("@variedad", SqlDbType.Int);
        //    cmd.Parameters["@variedad"].Value = cmb_variedad.SelectedValue.ToString();
        //    cn.Abrir();
        //    SqlDataAdapter adapter = new SqlDataAdapter(cmd);
        //    DataSet myds = new DataSet();
        //    adapter.Fill(myds);
        //    dataGridView4.DataSource = myds.Tables[0];
        //    cn.Cerrar();

        //}

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
            //System.Threading.Thread.Sleep(2000);
            //
            // FormLoading s = new FormLoading();
            //if(backgroundWorker1.IsBusy != true)
            //{
            //backgroundWorker1.RunWorkerAsync();
            //}
            Cursor = Cursors.WaitCursor;
            Grilla_total();
            Grilla_Calibre();
            Grilla_Categoria();
            tabControl1.Visible = true;
            Cursor = Cursors.Default;
            //dataGridView1.Columns[0].Visible = false;
            //dataGridView2.Columns[0].Visible = false;
            //dataGridView3.Columns[0].Visible = false;

            
           // s.ShowDialog();
           //   
       
  
            //
            //  pictureBox1.Visible = false;


        }

        public void Grilla_total()
        {
         
            SqlCommand cmd = new SqlCommand("spTraer_RealMasEstimado", ex.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Int);
            cmd.Parameters["@variedad"].Value = cmb_variedad.SelectedValue.ToString();

            ex.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView1.DataSource = myds.Tables[0];
            ex.Cerrar();
        }
        
        public void Grilla_Calibre()
        {
            SqlCommand cmd = new SqlCommand("spTraer_RealMasEstimadoCalibre", ex.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Int);
            cmd.Parameters["@variedad"].Value = cmb_variedad.SelectedValue.ToString();
            ex.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView2.DataSource = myds.Tables[0];
            ex.Cerrar();
        }
        public void Grilla_Categoria()
        {
            SqlCommand cmd = new SqlCommand("spTraer_RealMasEstimadoCategoria", ex.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@variedad", SqlDbType.Int);
            cmd.Parameters["@variedad"].Value = cmb_variedad.SelectedValue.ToString();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView3.DataSource = myds.Tables[0];
            ex.Cerrar();
        }

        private void dataGridView1_Scroll(object sender, ScrollEventArgs e)
        {
            dataGridView1.Columns[1].Frozen = true;
            dataGridView2.HorizontalScrollingOffset = e.NewValue;
            dataGridView3.HorizontalScrollingOffset = e.NewValue;
        }

        private void dataGridView2_Scroll(object sender, ScrollEventArgs e)
        {
            dataGridView2.Columns[1].Frozen = true;
            dataGridView1.HorizontalScrollingOffset = e.NewValue;
            dataGridView3.HorizontalScrollingOffset = e.NewValue;
        }

        private void dataGridView3_Scroll(object sender, ScrollEventArgs e)
        {
            dataGridView3.Columns[1].Frozen = true;
            dataGridView1.HorizontalScrollingOffset = e.NewValue;
            dataGridView2.HorizontalScrollingOffset = e.NewValue;
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
           try
            {

                int variedad = 2;
               
            }
            catch ( Exception ef)
            {
                MessageBox.Show(ef.Message);
             }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            pictureBox1.Visible = false;
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            try
            {
                if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString() == "")
                {
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = 0;
                }
            }
            catch
            { }
        }

        private void dataGridView2_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            try
            {
                if (dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString() == "")
                {
                    dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = 0;
                }
            }
            catch
            { }
        }

        private void dataGridView3_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            try
            {
                if (dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString() == "")
                {
                    dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = 0;
                }
            }
            catch
            { }
        }
        //public void EstimacionKG()
        //{
        //    SqlCommand cmd = new SqlCommand("spTraer_Estimacion", ex.getConexion());
        //    cmd.CommandType = CommandType.StoredProcedure;
        //    cmd.Parameters.Add("@variedad", SqlDbType.Int);
        //    cmd.Parameters["@variedad"].Value = cmb_variedad.SelectedValue.ToString();
        //    ex.Abrir();
        //    SqlDataAdapter adapter = new SqlDataAdapter(cmd);
        //    DataSet myds = new DataSet();
        //    adapter.Fill(myds);
        //    dataGridView1.DataSource = myds.Tables[0];
        //    ex.Cerrar();

        //}
        //public void EstimacionKG2()
        //{
        //    SqlCommand cmd = new SqlCommand("spTraer_Estimacion", ex.getConexion());
        //    cmd.CommandType = CommandType.StoredProcedure;
        //    cmd.Parameters.Add("@variedad", SqlDbType.Int);
        //    cmd.Parameters["@variedad"].Value = cmb_variedad.SelectedValue.ToString();
        //    ex.Abrir();
        //    SqlDataAdapter adapter = new SqlDataAdapter(cmd);
        //    DataSet myds = new DataSet();
        //    adapter.Fill(myds);
        //    dataGridView3.DataSource = myds.Tables[0];
        //    ex.Cerrar();

        //}

        //public void traerReal()
        //{
        //    try
        //    {
        //        SqlCommand cmd = new SqlCommand("spTraer_Estimacion2", ex.getConexion());
        //        cmd.CommandType = CommandType.StoredProcedure;
        //        cmd.Parameters.Add("@variedad", SqlDbType.VarChar, 25);
        //        cmd.Parameters["@variedad"].Value =cmb_variedad.SelectedValue.ToString();
        //        cn.Abrir();
        //        SqlDataAdapter adapter = new SqlDataAdapter(cmd);
        //        DataSet myds = new DataSet();
        //        adapter.Fill(myds);
        //        dataGridView2.DataSource = myds.Tables[0];
        //        cn.Cerrar();
        //    }
        //    catch (Exception e)
        //    {
        //        MessageBox.Show("Base de datos sin informacion :" + e, "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //    }

        //}

        //private void BtnFiltro_Click(object sender, EventArgs e)
        //{

        //    try
        //    {
        //        if ((cmb_variedad.SelectedValue.ToString() != "0") && (radioButton1.Checked == true))
        //        {
        //            Cursor = Cursors.WaitCursor;Estimacion_CalibrePORC();
        //            EstimacionKG();
        //            EstimacionKG2();

        //            traerReal();
        //            dataGridView1.Columns[0].Visible = false;
        //            dataGridView1.Columns[1].Visible = false;
        //            dataGridView2.Columns[1].Visible = false;
        //            dataGridView2.Columns[2].Visible = false;


        //            for (int x = 3; x < dataGridView3.ColumnCount - 1; x++)
        //            {
        //                double acum = 0;
        //                dataGridView3.Rows[dataGridView3.Rows.Count-1].Cells[x].Value = 0;
        //                for (int i = 0; i < dataGridView3.RowCount - 1; i++)
        //                {//dataGridView3.Rows[i].Cells[x].Value = 0;

        //                    try
        //                    {
        //                        dataGridView3.Rows[i].Cells[x].Value = Math.Round((Convert.ToDouble(dataGridView1.Rows[i].Cells[x].Value) - Convert.ToDouble(dataGridView2.Rows[i].Cells[x].Value)), 2);
        //                        acum = Math.Round((acum + (Convert.ToDouble(dataGridView1.Rows[i].Cells[x].Value) - Convert.ToDouble(dataGridView2.Rows[i].Cells[x].Value))), 2);
        //                        if (Convert.ToDouble(dataGridView3.Rows[i].Cells[x].Value.ToString()) < 0)
        //                        {
        //                            dataGridView3.Rows[i].Cells[x].Style.BackColor = Color.SteelBlue;
        //                            dataGridView3.Rows[i].Cells[x].Style.ForeColor = Color.White;
        //                            cantidad++;
        //                        }
        //                    }
        //                    catch { }

        //            }
        //                   dataGridView3.Rows[dataGridView3.Rows.Count - 1].Cells[x].Value = acum;
        //            }


        //            dataGridView3.Columns[0].Visible = false;
        //                   dataGridView3.Columns[1].Visible = false;
        //                   Cursor = Cursors.Default;
        //                   tabControl1.Visible = true;
        //        }
        //        if(cantidad > 0)
        //        {
        //            MessageBox.Show("Necesita actualizar planillas de estimacion" + Environment.NewLine + "Debido a :" + Environment.NewLine +
        //                " " + Environment.NewLine
        //                + " * Se encuentran productos procesados mayor a su estimación.", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //        }

        //    }           

        //    catch (Exception x)
        //    {
        //        MessageBox.Show("se produjo un error al tratar de cargar datos  " + x.ToString());
        //    }  
        //}

        //private void dataGridView2_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        //{

        //    if (dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString() == "")
        //    {
        //        dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = 0;
        //    }

        //}

        //        private void dataGridView3_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        //        {
        //            try
        //            {
        //                if (Convert.ToDouble(dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()) < 0)
        //                {
        //                    dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.SteelBlue;
        //                    dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.ForeColor = Color.White;


        //                }
        //                if ((Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()) == 0) && ( (dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()=="")||(Convert.ToDouble(dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()) == 0) ) )//(Convert.ToDouble(dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()) == 0)||
        //                {
        //                    dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.OrangeRed;
        //                    dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.ForeColor = Color.White;
        //                }
        //                if ((Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()) > 0) && ((dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString() == "") || (Convert.ToDouble(dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()) == 0)))
        //                {
        //                    dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Olive;
        //                    dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.ForeColor = Color.White;
        //                }
        //            }
        //            catch { }
        //        }

        //        private void dataGridView2_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        //        {
        //  if (dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString() == "")
        //            {
        //                dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = 0;
        //            }
        //}

        //        private void dataGridView2_CellEndEdit(object sender, DataGridViewCellEventArgs e) 
        //        {
        //            if (dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString() == "")
        //            {
        //                dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = 0;
        //            }
        //        }
    }  
}
