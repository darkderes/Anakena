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
        conexionERP cn = new conexionERP();
        conexion ex = new conexion();
        public FormRealidad()
        {
            InitializeComponent();
        }

        private void FormRealidad_Load(object sender, EventArgs e)
        {
            
            CmbVariedad();
            traerDistintosProductores();
            rellenar_productores();
            radioButton1.Checked = true;
            CmbTipo.SelectedIndex = 0;
    

        }
        public void traerReal()
        {
            try
            {
            SqlCommand cmd = new SqlCommand("sistema_realidad_all", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView4.DataSource = myds.Tables[0];
            cn.Cerrar();
            }
            catch(Exception e)
            { MessageBox.Show("Base de datos sin informacion :"+e, "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }
        public void traer_estimacion(string productor, string variedad)
        {
            SqlCommand cmd = new SqlCommand("spTraer_Estimacion2", ex.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@productor", SqlDbType.VarChar, 25);
            cmd.Parameters["@productor"].Value = productor;
            cmd.Parameters.Add("@variedad", SqlDbType.VarChar, 25);
            cmd.Parameters["@variedad"].Value = variedad;
            ex.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView3.DataSource = myds.Tables[0];
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

        public Double agrupar(string categoria,string productor,string calibre,string variedad)
        {
            Double acum36 = 0;
            for (int i = 0;i < dataGridView4.RowCount; i++)
            {
              try
                {
                    if ((dataGridView4.Rows[i].Cells["Categorias"].Value.ToString().Trim() == categoria.Trim()) && (dataGridView4.Rows[i].Cells["Cod_Productor"].Value.ToString() == productor) && categoria != null && (dataGridView4.Rows[i].Cells["Cod_Variedad"].Value.ToString() == variedad))
                {
                       
                         acum36 = acum36 + Convert.ToDouble(dataGridView4.Rows[i].Cells[calibre].Value.ToString()); 
                }
                else if ((categoria == " ") && (dataGridView4.Rows[i].Cells["Cod_Productor"].Value.ToString() == productor) && (dataGridView4.Rows[i].Cells["Cod_Variedad"].Value.ToString() == variedad))
                    {
                       
                        acum36 = acum36 + Convert.ToDouble(dataGridView4.Rows[i].Cells[calibre].Value.ToString());
                    }
                }
                catch { }
              
            }  return acum36;
       
        }
         public Double agrupar_categoria(string categoria, string productor, string variedad)
        {
           Double acum = 0;
            for(int i = 1; i< dataGridView4.RowCount;i++)
            {
                try
                    {
                    
                    if ((dataGridView4.Rows[i].Cells["Categorias"].Value.ToString().Trim() == categoria.Trim()) && (dataGridView4.Rows[i].Cells["Cod_Productor"].Value.ToString() == productor)  && (dataGridView4.Rows[i].Cells["Cod_Variedad"].Value.ToString() == variedad))
                    {
                        
                        acum = acum +( Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value.ToString())+ Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value.ToString())+ Convert.ToDouble(dataGridView4.Rows[i].Cells[4].Value.ToString()) + Convert.ToDouble(dataGridView4.Rows[i].Cells[5].Value.ToString())+ Convert.ToDouble(dataGridView4.Rows[i].Cells[6].Value.ToString())+ Convert.ToDouble(dataGridView4.Rows[i].Cells[7].Value.ToString()) + Convert.ToDouble(dataGridView4.Rows[i].Cells[8].Value.ToString())+ Convert.ToDouble(dataGridView4.Rows[i].Cells[9].Value.ToString()));
                      //  MessageBox.Show(acum.ToString());
                    }

                    
                }catch { }
            }
            return acum;
        }
        public Double agrupar_calibre(string calibre, string productor, string variedad)
        {
            Double acum = 0;
            for (int i = 1; i < dataGridView4.RowCount; i++)
            {
                try
                {

                    if ((dataGridView4.Rows[i].Cells["Cod_Productor"].Value.ToString() == productor) && (dataGridView4.Rows[i].Cells["Cod_Variedad"].Value.ToString() == variedad))
                    {
                        acum = acum + Convert.ToDouble(dataGridView4.Rows[i].Cells[calibre].Value.ToString()); //+ Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value.ToString()) + Convert.ToDouble(dataGridView4.Rows[i].Cells[4].Value.ToString()) + Convert.ToDouble(dataGridView4.Rows[i].Cells[5].Value.ToString())+ Convert.ToDouble(dataGridView4.Rows[i].Cells[6].Value.ToString())+ Convert.ToDouble(dataGridView4.Rows[i].Cells[7].Value.ToString()) + Convert.ToDouble(dataGridView4.Rows[i].Cells[8].Value.ToString())+ Convert.ToDouble(dataGridView4.Rows[i].Cells[9].Value.ToString());
                                                                                                         // MessageBox.Show(acum.ToString());
                    }
                }
                catch { }
            }
            return acum;
        }

        public void traerDistintosProductores()
        {

            SqlCommand cmd = new SqlCommand("spTraerDistintosProductores", ex.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            ex.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView1.DataSource = myds.Tables[0];
            ex.Cerrar();

        }
    


        public void rellenar_productores()
        {
            DataGridViewTextBoxColumn col12 = new DataGridViewTextBoxColumn();
            col12.HeaderText = "Productores";
            dataGridView2.Columns.Add(col12);
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn();
                col.HeaderText = dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                col.Name = dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                dataGridView2.Columns.Add(col);

            }
        }

        public void rellenar_productores_calidad()
        {
            DataGridViewTextBoxColumn col12 = new DataGridViewTextBoxColumn();
            col12.HeaderText = "Calidad";
            dataGridView5.Columns.Add(col12);
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn();
                col.HeaderText = dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                col.Name = dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                dataGridView5.Columns.Add(col);

            }
        }
        public void rellenar_productores_calibre()
        {
            DataGridViewTextBoxColumn col12 = new DataGridViewTextBoxColumn();
            col12.HeaderText = "Calibre";
            dataGridView6.Columns.Add(col12);
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn();
                col.HeaderText = dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                col.Name = dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                dataGridView6.Columns.Add(col);

            }
        }
        private void BtnFiltro_Click(object sender, EventArgs e)
        {
            limpiar();            
            dataGridView2.DataSource = null;
            dataGridView2.Rows.Clear();
            pictureBox2.Visible = false;
            pictureBox1.Visible = true;
            pictureBox1.Show();
            pictureBox1.Refresh();
            // label2.Visible = true;
            pBar1.Visible = true;
            if(radioButton1.Checked == true)
            { 
         
            traerReal();
            traer_estimacion(dataGridView1.Rows[0].Cells[0].Value.ToString(), cmb_variedad.Text);
            rellenar_serr(dataGridView1.Rows[0].Cells[0].Value.ToString());
            pBar1.Minimum = 1;
            pBar1.Maximum = dataGridView2.RowCount * dataGridView2.ColumnCount;
            pBar1.Value = 1;
            pBar1.Step = 1;
       
            proceso();
            dataGridView2.Visible = true;
            pBar1.Visible = false;
                    
            }
           else
            {
              
                rellenar_productores_calidad();
                rellenar_productores_calibre();
                dataGridView5.Rows.Insert(0,"EXTRA");
                dataGridView5.Rows.Insert(1, "CAT I");
                dataGridView5.Rows.Insert(2, "CAT II");
                dataGridView5.Rows.Insert(3, "CAT III");
                dataGridView5.Rows.Insert(4, "BAJO CAT III");
                dataGridView5.Rows.Insert(5, "Total");

                //datagriedview 6
                dataGridView6.Rows.Insert(0, "36+");
                dataGridView6.Rows.Insert(1, "34-36");
                dataGridView6.Rows.Insert(2, "32-34");
                dataGridView6.Rows.Insert(3, "30-32");
                dataGridView6.Rows.Insert(4, "28-30");
                dataGridView6.Rows.Insert(5, "PRECALIBRE CALIBRADO");
                dataGridView6.Rows.Insert(6, "VANAS CALIBRADO");
                dataGridView6.Rows.Insert(7, "MERMA CALIBRADO");
                dataGridView6.Rows.Insert(8, "Total");
                dataGridView6.Rows[dataGridView6.Rows.Count - 1].DefaultCellStyle.BackColor = Color.SteelBlue;
                dataGridView6.Rows[dataGridView6.Rows.Count - 1].DefaultCellStyle.ForeColor = Color.White;
                dataGridView5.Rows[dataGridView5.Rows.Count - 1].DefaultCellStyle.BackColor = Color.SteelBlue;
                dataGridView5.Rows[dataGridView5.Rows.Count - 1].DefaultCellStyle.ForeColor = Color.White;
                traerReal();
                pBar1.Minimum = 1;
                pBar1.Maximum = (dataGridView5.RowCount * dataGridView5.ColumnCount)+(dataGridView6.RowCount * dataGridView6.ColumnCount);
                pBar1.Value = 1;
                pBar1.Step = 1;
                for (int i = 1;i<dataGridView5.Columns.Count;i++)
                {
                    double acum = 0;
                    for(int x = 0;x<dataGridView5.Rows.Count;x++)
                    {
                        double valor = agrupar_categoria(dataGridView5.Rows[x].Cells[0].Value.ToString(), dataGridView5.Columns[i].Name.ToString(), cmb_variedad.SelectedValue.ToString());
                        acum = acum + valor;
                        dataGridView5.Rows[x].Cells[i].Value = valor;
                        pBar1.PerformStep();
                    }
                    dataGridView5.Rows[dataGridView5.RowCount - 1].Cells[i].Value = acum;
                }
                for (int i = 1; i < dataGridView6.Columns.Count; i++)
                {
                    double acum = 0;
                    for (int x = 0; x < dataGridView6.Rows.Count; x++)
                    {
                        double valor = agrupar_calibre(dataGridView6.Rows[x].Cells[0].Value.ToString(), dataGridView6.Columns[i].Name.ToString(), cmb_variedad.SelectedValue.ToString());
                        acum = acum + valor;
                        dataGridView6.Rows[x].Cells[i].Value = valor;
                        pBar1.PerformStep();
                    }
                    dataGridView6.Rows[dataGridView6.RowCount - 1].Cells[i].Value = acum;
                }
                Btn_Excel.Visible = true;
                if (CmbTipo.Text == "%" && radioButton2.Checked == true)
                {
                    for (int u = 1; u < dataGridView5.Columns.Count; u++)
                    {
                        double total = Convert.ToDouble((dataGridView5.Rows[dataGridView5.Rows.Count - 1].Cells[u].Value.ToString()));
                        for (int y = 0; y < dataGridView5.Rows.Count; y++)
                        {
                            if (total != 0)
                            {
                                dataGridView5.Rows[y].Cells[u].Value = (Convert.ToDouble(dataGridView5.Rows[y].Cells[u].Value.ToString()) / total) * 100;
                            }
                        }

                    }
                }
                if (CmbTipo.Text == "%" && radioButton2.Checked == true)
                {
                    for (int u = 1; u < dataGridView6.Columns.Count; u++)
                    {
                        double total = Convert.ToDouble((dataGridView6.Rows[dataGridView6.Rows.Count - 1].Cells[u].Value.ToString()));
                        for (int y = 0; y < dataGridView6.Rows.Count; y++)
                        {
                            if (total != 0)
                            {
                                dataGridView6.Rows[y].Cells[u].Value = (Convert.ToDouble(dataGridView6.Rows[y].Cells[u].Value.ToString()) / total) * 100;
                            }
                        }

                    }
                }
                dataGridView2.Visible = false;
                dataGridView5.Visible = true;
                dataGridView6.Visible = true;
            }
            pictureBox2.Visible = true;
           pictureBox1.Visible = false;
        }
        public void proceso()
        {
            for (int i = 1; i < dataGridView2.Columns.Count; i++)
            {
                if(radioButton1.Checked == true)
                { 
                double acum = 0;
                for (int x = 0; x < dataGridView2.RowCount; x++)
                {

                    string y = dataGridView2.Rows[x].Cells[0].Value.ToString().Trim();
                    string[] valor = y.Split('/');
                    try
                    {
                        string calibre = valor[0];
                        string calidad = valor[1];
                        double valores = agrupar(calidad, dataGridView2.Columns[i].Name.ToString(), calibre, cmb_variedad.SelectedValue.ToString());
                        dataGridView2.Rows[x].Cells[i].Value = valores;
                        acum = acum + valores;
                        pBar1.PerformStep();
                        label2.Text = "Proceso :" + pBar1.Value.ToString() + "/" + (dataGridView2.RowCount * dataGridView2.ColumnCount).ToString();
                        
                    }
                    catch
                    {
                        string cal = valor[0];
                        double valores = agrupar(" ", dataGridView2.Columns[i].Name.ToString(), cal, cmb_variedad.SelectedValue.ToString());
                        dataGridView2.Rows[x].Cells[i].Value = valores;
                        acum = acum + valores;
                        pBar1.PerformStep();
                        label2.Text = "Proceso :" + pBar1.Value.ToString() + "/" + (dataGridView2.RowCount * dataGridView2.ColumnCount).ToString();
                    }
                }
                 dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[i].Value = acum;
               
                }
             
            }
            if (CmbTipo.Text == "%" && radioButton1.Checked == true)
            {
                for (int u = 1; u < dataGridView2.Columns.Count; u++)
                {
                    double total = Convert.ToDouble((dataGridView2.Rows[dataGridView2.Rows.Count-1].Cells[u].Value.ToString()));
                    for(int y = 0;y<dataGridView2.Rows.Count; y++)
                    {
                       if(total != 0)
                        { 
                        dataGridView2.Rows[y].Cells[u].Value = (Convert.ToDouble(dataGridView2.Rows[y].Cells[u].Value.ToString()) / total) * 100;
                    }}
                   
                }
            }
            Btn_Excel.Visible = true;
        }

 
        public void rellenar_serr(string columna)
        {
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                if (dataGridView2.Columns[0].HeaderText == "Productores")
                {
                    dataGridView2.Rows.Insert(i, dataGridView3.Rows[i].Cells[0].Value.ToString());
                }
            }
            dataGridView2.Rows.Insert(dataGridView2.Rows.Count, "TOTAL");
            dataGridView2.Rows[dataGridView2.Rows.Count - 1].DefaultCellStyle.BackColor = Color.SteelBlue;
            dataGridView2.Rows[dataGridView2.Rows.Count - 1].DefaultCellStyle.ForeColor = Color.White;
        }
        public void exporta_a_excel()
        {
            ApplicationClass class2 = new ApplicationClass();
            class2.Application.Workbooks.Add(true);
            int num = 0;
            pBar1.Minimum = 1;
            pBar1.Maximum = dataGridView2.RowCount * dataGridView2.ColumnCount;
            pBar1.Value = 1;
            pBar1.Step = 1;

            foreach (DataGridViewColumn column in this.dataGridView2.Columns)
            {
                class2.Cells[1, 1] = cmb_variedad.Text.Trim();
                num++;
                class2.Cells[1, num] = column.Name;
                class2.get_Range("A1", "GB1").Interior.ColorIndex = 9;
                class2.get_Range("A1", "GB1").Font.ColorIndex = 2;
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
                    label2.Text = "Proceso :" + pBar1.Value.ToString() + "/" + (dataGridView2.RowCount * dataGridView2.ColumnCount).ToString();
                }
            }
            class2.Visible = true;
            ((Worksheet)class2.ActiveSheet).Name = cmb_variedad.Text.Trim();
            ((Worksheet)class2.ActiveSheet).Activate();

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


        private void Btn_Excel_Click(object sender, EventArgs e)
        {
            if (dataGridView2.Visible == true)
            {
                dataGridView2.Visible = false;
            exporta_a_excel();
            }
           
            else
            {
                pictureBox1.Visible = true;
                dataGridView5.Visible = false;
                dataGridView6.Visible = false;

                    exporta_a_excel_categoria();
            }
            pBar1.Visible = false; Btn_Excel.Visible = false;
        }

        private void dataGridView4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView2_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (CmbTipo.Text == "%")
            {
                e.CellStyle.Format = "0.00\\% ";
            }
            else
            {
                e.CellStyle.Format = "0.00\\";
            }
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dataGridView5_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (CmbTipo.Text == "%")
            {
                e.CellStyle.Format = "0.00\\% ";
            }
            else
            {
                e.CellStyle.Format = "0.00\\";
            }
        }

        private void dataGridView6_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (CmbTipo.Text == "%")
            {
                e.CellStyle.Format = "0.00\\% ";
            }
            else
            {
                e.CellStyle.Format = "0.00\\";
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            limpiar();
        }
        public void limpiar()
        {
            dataGridView2.Visible = false;
            dataGridView2.Rows.Clear();
            dataGridView2.DataSource = null;

            dataGridView5.Visible = false;
            dataGridView5.Rows.Clear();
            dataGridView5.DataSource = null;

            dataGridView6.Visible = false;
            dataGridView6.Rows.Clear();
            dataGridView6.DataSource = null;

            Btn_Excel.Visible = false;
            label2.Visible = false;
            pBar1.Visible = false;
            btnCancelar.Visible = false;
        }
    }
    
}
