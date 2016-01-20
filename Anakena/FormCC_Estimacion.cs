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
    public partial class FormCC_Estimacion : Form
    {
        conexion cn = new conexion();
        double acum_total = 0;
        public FormCC_Estimacion()
        {
            InitializeComponent();
        }

        private void FormCC_Estimacion_Load(object sender, EventArgs e)
        {
            CmbVariedad();
            traerDistintosProductores();
            rellenar_productores();
       
        }
        public void traerDistintosProductores()
        {

            SqlCommand cmd = new SqlCommand("spTraerDistintosProductores", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView1.DataSource = myds.Tables[0];
            cn.Cerrar();

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
                MessageBox.Show("Ocurrio un problema al cargar variedades", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cn.Cerrar();
            }

        }
        public void traer_estimacion(string productor, string variedad)
        {
            SqlCommand cmd = new SqlCommand("spTraer_Estimacion", cn.getConexion());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@productor", SqlDbType.VarChar, 25);
            cmd.Parameters["@productor"].Value = productor;
            cmd.Parameters.Add("@variedad", SqlDbType.VarChar, 25);
            cmd.Parameters["@variedad"].Value = variedad;
            cn.Abrir();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet myds = new DataSet();
            adapter.Fill(myds);
            dataGridView3.DataSource = myds.Tables[0];
            cn.Cerrar();

        }

        private void BtnFiltro_Click(object sender, EventArgs e)
        {
            if (cmb_variedad.SelectedValue.ToString() != "0")
            {
                dataGridView2.Rows.Clear();
                dataGridView2.DataSource = null;
                traer_estimacion(dataGridView1.Rows[0].Cells[0].Value.ToString(), cmb_variedad.Text);
                rellenar_serr(dataGridView1.Rows[0].Cells[0].Value.ToString());
                relleno_Serr();
                relleno_grilla();
              /*  tabControl1.Visible = true;
                Btn_Excel.Visible = true;
                tabPage1.Text = cmb_variedad.Text;
                pBar1.Visible = false;
                label2.Visible = false;*/
            }
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
        public void relleno_Serr()
        {
          
         /*   pBar1.Minimum = 1;
            pBar1.Maximum = dataGridView2.RowCount * dataGridView2.ColumnCount;
            pBar1.Value = 1;
            pBar1.Step = 1;
            label2.Visible = true;
            pBar1.Visible = true;*/
            for (int i = 1; i < dataGridView2.ColumnCount; i++)
            {
                double acum = 0;
                traer_estimacion(dataGridView2.Columns[i].Name, cmb_variedad.Text);
                for (int x = 0; x < dataGridView3.Rows.Count; x++)
                {
                    dataGridView2.Rows[x].Cells[i].Value = dataGridView3.Rows[x].Cells[3].Value.ToString();
                    acum = acum + Convert.ToDouble(dataGridView2.Rows[x].Cells[i].Value.ToString());
             //       pBar1.PerformStep();
             //       label2.Text = "Proceso :" + pBar1.Value.ToString() + "/" + (dataGridView2.RowCount * dataGridView2.ColumnCount).ToString();
                }
                dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[i].Value = acum;
                acum_total = acum_total + acum;
              

            }
        }
        public void relleno_grilla()
        {
            pBar1.Minimum = 1;
            pBar1.Maximum = (dataGridView2.RowCount * dataGridView2.ColumnCount) ;
            pBar1.Value = 1;
            pBar1.Step = 1;
            DataGridViewTextBoxColumn Column1 =
               new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn Column2 =
               new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn Column3 =
            new DataGridViewTextBoxColumn();
            Column1.Name = "Detalle";
            Column1.HeaderText = "Calibre/Categoria";

            Column2.Name = "Valor";
            Column2.HeaderText = "Valor";

            Column3.Name = "Porcentaje";
            Column3.HeaderText = "Porcentaje";


            dataGridView4.Columns.Insert(0, Column1);
            dataGridView4.Columns.Insert(1, Column2);
            dataGridView4.Columns.Insert(2, Column3);

            double acumTotal = 0;
            double acumPorcentaje = 0;
            double r = Convert.ToDouble(dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[3].Value.ToString());
            for (int x = 0; x<dataGridView2.Rows.Count-1;x++)
            {
            double acum = 0;
           
                dataGridView4.Rows.Insert(x, dataGridView2.Rows[x].Cells[0].Value.ToString());
  
            for (int i = 0; i < dataGridView2.Columns.Count-1; i++)
            {
                    try
                    {
                        if(i >0)
                        {
                        acum = acum +Convert.ToDouble(dataGridView2.Rows[x].Cells[i].Value.ToString());
                        }                  
                    }
                    catch { }
                    pBar1.PerformStep();

                }
                dataGridView4.Rows[x].Cells["Valor"].Value = acum;
               string porcentaje =  ((acum * 100) / acum_total).ToString("n2");
                dataGridView4.Rows[x].Cells["Porcentaje"].Value = porcentaje;
                
                // MessageBox.Show(acum.ToString());
                pBar1.PerformStep();
                acumTotal = acumTotal + acum;
                acumPorcentaje = acumPorcentaje + Convert.ToDouble(porcentaje);            }
            dataGridView4.Rows.Insert(dataGridView4.Rows.Count, "TOTAL");
            dataGridView4.Rows[dataGridView4.Rows.Count-1].Cells[1].Value = acumTotal.ToString();
            pBar1.Visible = false;
            dataGridView4.Rows[dataGridView4.Rows.Count - 1].Cells[2].Value =  Math.Round(acumPorcentaje).ToString();
         
        }
        public void exporta_a_excel()
        {
            ApplicationClass class2 = new ApplicationClass();
            class2.Application.Workbooks.Add(true);
            int num = 0;
            pBar1.Minimum = 1;
            pBar1.Maximum = dataGridView4.RowCount * dataGridView4.ColumnCount;
            pBar1.Value = 1;
            pBar1.Step = 1;

            foreach (DataGridViewColumn column in this.dataGridView4.Columns)
            {
                class2.Cells[1, 1] = cmb_variedad.Text.Trim();
                num++;
                class2.Cells[1, num] = column.Name;
                class2.get_Range("A1", "C1").Interior.ColorIndex = 9;
                class2.get_Range("A1", "C1").Font.ColorIndex = 2;
            }
            int num2 = 0;
            foreach (DataGridViewRow row in (IEnumerable)this.dataGridView4.Rows)
            {
                num2++;
                num = 0;
                foreach (DataGridViewColumn column2 in this.dataGridView4.Columns)
                {
                    num++;
                    class2.Cells[num2 + 1, num] = row.Cells[column2.Name].Value;
                    pBar1.PerformStep();
                  //  label2.Text = "Proceso :" + pBar1.Value.ToString() + "/" + (dataGridView2.RowCount * dataGridView2.ColumnCount).ToString();
                }
            }
            class2.Visible = true;
            ((Worksheet)class2.ActiveSheet).Name = cmb_variedad.Text.Trim();
            ((Worksheet)class2.ActiveSheet).Activate();

        }

        private void Btn_Excel_Click(object sender, EventArgs e)
        {
            exporta_a_excel();
        }
    }
}
