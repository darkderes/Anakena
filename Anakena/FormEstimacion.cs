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
           traerDistintosProductores();
            tabControl1.Visible = false;
            rellenar();
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
        public void traer_estimacion(string productor,string variedad)
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
        public void rellenar_serr(string columna)
        {
          
            for(int i = 0;i<dataGridView3.Rows.Count;i++)
            {
                if (dataGridView2.Columns[0].HeaderText == "Productores")
                {
                    dataGridView2.Rows.Insert(i, dataGridView3.Rows[i].Cells[0].Value.ToString());
                }         
            }
            dataGridView2.Rows.Insert(dataGridView2.Rows.Count,"TOTAL");
            dataGridView2.Rows[dataGridView2.Rows.Count-1].DefaultCellStyle.BackColor = Color.SteelBlue;
            dataGridView2.Rows[dataGridView2.Rows.Count - 1].DefaultCellStyle.ForeColor = Color.White;

        }
        public void relleno_Serr()
        {
            
            pBar1.Minimum = 1;
            pBar1.Maximum = dataGridView2.RowCount * dataGridView2.ColumnCount;
            pBar1.Value = 1;
            pBar1.Step = 1;
            label2.Visible = true;
            pBar1.Visible = true;
            for (int i = 1; i < dataGridView2.ColumnCount; i++)
            {
                double acum = 0;
               traer_estimacion(dataGridView2.Columns[i].Name, cmb_variedad.Text);
                for (int x = 0; x < dataGridView3.Rows.Count; x++)
                {
                   dataGridView2.Rows[x].Cells[i].Value = dataGridView3.Rows[x].Cells[3].Value.ToString();
                   acum = acum + Convert.ToDouble(dataGridView2.Rows[x].Cells[i].Value.ToString());
                   pBar1.PerformStep();
                   label2.Text = "Proceso :" + pBar1.Value.ToString() + "/" + (dataGridView2.RowCount * dataGridView2.ColumnCount).ToString();
                }
                dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[i].Value = acum;
               
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
                MessageBox.Show("Ocurrio un problema al cargar variedades","Anakena",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
            finally
            {
                cn.Cerrar();
            }
        }
  
        public void  rellenar()
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
           try
            {
                if(cmb_variedad.SelectedValue.ToString() != "0")
                {
                dataGridView2.Rows.Clear();
                dataGridView2.DataSource = null;
                traer_estimacion(dataGridView1.Rows[0].Cells[0].Value.ToString(), cmb_variedad.Text);
                rellenar_serr(dataGridView1.Rows[0].Cells[0].Value.ToString());
                relleno_Serr();
                tabControl1.Visible = true;
                Btn_Excel.Visible = true;
                tabPage1.Text = cmb_variedad.Text;
                pBar1.Visible = false;
                label2.Visible = false;
                }
                else
                {

                    for (int i = 1; i < 3; i++)
                    {
                        dataGridView2.Rows.Clear();
                        dataGridView2.DataSource = null;
                        cmb_variedad.SelectedIndex = i;
                        traer_estimacion(dataGridView1.Rows[0].Cells[0].Value.ToString(), cmb_variedad.Text);
                        rellenar_serr(dataGridView1.Rows[0].Cells[0].Value.ToString());
                        relleno_Serr();
                        tabControl1.Visible = true;
                        // Btn_Excel.Visible = true;
                        tabPage1.Text = cmb_variedad.Text;
                        exporta_a_excel();
                    }
                }
            }
            catch
            {
                tabControl1.Visible = false;
                Btn_Excel.Visible = false;
                MessageBox.Show("Ocurrio un problema al tratar de mostrar estimulacion de variedad", "Anakena", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
                
        }
        public void exporta_a_excel()
        {
            ApplicationClass class2 = new ApplicationClass();
            class2.Application.Workbooks.Add(true);
            int num = 0;
            pBar1.Minimum = 1;
            pBar1.Maximum = dataGridView2.RowCount * dataGridView2.ColumnCount ;
            pBar1.Value = 1;
            pBar1.Step = 1;
       
            foreach (DataGridViewColumn column in this.dataGridView2.Columns)
            {
              class2.Cells[1, 1] = cmb_variedad.Text.Trim();
              num++;
              class2.Cells[1, num] = column.Name;
              class2.get_Range("A1", "FK1").Interior.ColorIndex = 9;
              class2.get_Range("A1", "FK1").Font.ColorIndex = 2;   
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
            class2.Visible = true;
            ((Worksheet)class2.ActiveSheet).Name = cmb_variedad.Text.Trim();
            ((Worksheet)class2.ActiveSheet).Activate();
             
        }
    }
}
