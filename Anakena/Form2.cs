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
    public partial class Form2 : Form
    {
        conexion cn = new conexion();
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {         
            traerDistintosProductores();
            rellenar();
            tabControl1.SelectedIndex = 0;
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
           
        }
        public void rellenar_CHARLERD()
        {

            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                if (dataGridView4.Columns[0].HeaderText == "Productores")
                {
                    dataGridView4.Rows.Insert(i, dataGridView3.Rows[i].Cells[0].Value.ToString());
                }
            }

        }
        public void rellenar_howard()
        {

            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                if (dataGridView5.Columns[0].HeaderText == "Productores")
                {
                    dataGridView5.Rows.Insert(i, dataGridView3.Rows[i].Cells[0].Value.ToString());
                }
            }

        }
        public void rellenar_tulare()
        {

            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                if (dataGridView6.Columns[0].HeaderText == "Productores")
                {
                    dataGridView6.Rows.Insert(i, dataGridView3.Rows[i].Cells[0].Value.ToString());
                }
            }

        }
        public void rellenar_sundland()
        {

            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                if (dataGridView7.Columns[0].HeaderText == "Productores")
                {
                    dataGridView7.Rows.Insert(i, dataGridView3.Rows[i].Cells[0].Value.ToString());
                }
            }

        }
        public void rellenar_hartley()
        {

            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                if (dataGridView8.Columns[0].HeaderText == "Productores")
                {
                    dataGridView8.Rows.Insert(i, dataGridView3.Rows[i].Cells[0].Value.ToString());
                }
            }

        }
        public void rellenar_semilla()
        {

            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                if (dataGridView9.Columns[0].HeaderText == "Productores")
                {
                    dataGridView9.Rows.Insert(i, dataGridView3.Rows[i].Cells[0].Value.ToString());
                }
            }

        }
        public void rellenar_vina()
        {

            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                if (dataGridView10.Columns[0].HeaderText == "Productores")
                {
                    dataGridView10.Rows.Insert(i, dataGridView3.Rows[i].Cells[0].Value.ToString());
                }
            }

        }
       
        public void rellenar_vina_mejorada()
        {

            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                if (dataGridView11.Columns[0].HeaderText == "Productores")
                {
                    dataGridView11.Rows.Insert(i, dataGridView3.Rows[i].Cells[0].Value.ToString());
                }
            }

        }
        public void rellenar_franquette()
        {

            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                if (dataGridView12.Columns[0].HeaderText == "Productores")
                {
                    dataGridView12.Rows.Insert(i, dataGridView3.Rows[i].Cells[0].Value.ToString());
                }
            }

        }
        public void rellenar_dctedespelonado()
        {

            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                if (dataGridView13.Columns[0].HeaderText == "Productores")
                {
                    dataGridView13.Rows.Insert(i, dataGridView3.Rows[i].Cells[0].Value.ToString());
                }
            }

        }

        public void rellenar_dctequebrado()
        {
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                if (dataGridView14.Columns[0].HeaderText == "Productores")
                {
                    dataGridView14.Rows.Insert(i, dataGridView3.Rows[i].Cells[0].Value.ToString());
                }
            }

        }
        public void relleno_Serr()
        {
            for(int i = 1; i<dataGridView2.ColumnCount; i++)
            {
                traer_estimacion(dataGridView2.Columns[i].Name, "SERR");
                for (int x = 0; x < dataGridView3.Rows.Count;x++)
                {
                    dataGridView2.Rows[x].Cells[i].Value = dataGridView3.Rows[x].Cells[3].Value.ToString();
                  
                }          
            }
        }
        public void relleno_chandler()
        {
            for (int i = 1; i < dataGridView4.ColumnCount; i++)
            {
                traer_estimacion(dataGridView4.Columns[i].Name, "CHANDLER");
                for (int x = 0; x < dataGridView3.Rows.Count; x++)
                {
                    dataGridView4.Rows[x].Cells[i].Value = dataGridView3.Rows[x].Cells[3].Value.ToString();

                }
            }
        }
        public void relleno_howard()
        {
            for (int i = 1; i < dataGridView5.ColumnCount; i++)
            {
                traer_estimacion(dataGridView5.Columns[i].Name, "HOWARD");
                for (int x = 0; x < dataGridView3.Rows.Count; x++)
                {
                    dataGridView5.Rows[x].Cells[i].Value = dataGridView3.Rows[x].Cells[3].Value.ToString();

                }
            }
        }
        public void relleno_tulare()
        {
            for (int i = 1; i < dataGridView6.ColumnCount; i++)
            {
                traer_estimacion(dataGridView6.Columns[i].Name, "TULARE");
                for (int x = 0; x < dataGridView3.Rows.Count; x++)
                {
                    dataGridView6.Rows[x].Cells[i].Value = dataGridView3.Rows[x].Cells[3].Value.ToString();

                }
            }
        }
        public void relleno_sundland()
        {
            for (int i = 1; i < dataGridView7.ColumnCount; i++)
            {
                traer_estimacion(dataGridView7.Columns[i].Name, "SUNLAND");
                for (int x = 0; x < dataGridView3.Rows.Count; x++)
                {
                    dataGridView7.Rows[x].Cells[i].Value = dataGridView3.Rows[x].Cells[3].Value.ToString();

                }
            }
        }
        public void relleno_hartley()
        {
            for (int i = 1; i < dataGridView8.ColumnCount; i++)
            {
                traer_estimacion(dataGridView8.Columns[i].Name, "HARTLEY");
                for (int x = 0; x < dataGridView3.Rows.Count; x++)
                {
                    dataGridView8.Rows[x].Cells[i].Value = dataGridView3.Rows[x].Cells[3].Value.ToString();

                }
            }
        }
        public void relleno_semilla()
        {
            for (int i = 1; i < dataGridView9.ColumnCount; i++)
            {
                traer_estimacion(dataGridView9.Columns[i].Name, "SEMILLA");
                for (int x = 0; x < dataGridView3.Rows.Count; x++)
                {
                    dataGridView9.Rows[x].Cells[i].Value = dataGridView3.Rows[x].Cells[3].Value.ToString();

                }
            }
        }
        public void relleno_vina()
        {
            for (int i = 1; i < dataGridView10.ColumnCount; i++)
            {
                traer_estimacion(dataGridView10.Columns[i].Name, "VINA");
                for (int x = 0; x < dataGridView3.Rows.Count; x++)
                {
                    dataGridView10.Rows[x].Cells[i].Value = dataGridView3.Rows[x].Cells[3].Value.ToString();

                }
            }
        }
        
        public void relleno_vina_mejorada()
        {
            for (int i = 1; i < dataGridView11.ColumnCount; i++)
            {
                traer_estimacion(dataGridView11.Columns[i].Name, "VINA MEJORADA");
                for (int x = 0; x < dataGridView3.Rows.Count; x++)
                {
                    dataGridView11.Rows[x].Cells[i].Value = dataGridView3.Rows[x].Cells[3].Value.ToString();

                }
            }
        }
        public void relleno_franquette()
        {
            for (int i = 1; i < dataGridView12.ColumnCount; i++)
            {
                traer_estimacion(dataGridView12.Columns[i].Name, "FRANQUETTE");
                for (int x = 0; x < dataGridView3.Rows.Count; x++)
                {
                    dataGridView12.Rows[x].Cells[i].Value = dataGridView3.Rows[x].Cells[3].Value.ToString();

                }
            }
        }
        public void relleno_dctedespelonado()
        {
            for (int i = 1; i < dataGridView13.ColumnCount; i++)
            {
                traer_estimacion(dataGridView13.Columns[i].Name, "DCTE DESPELONADO");
                for (int x = 0; x < dataGridView3.Rows.Count; x++)
                {
                    dataGridView13.Rows[x].Cells[i].Value = dataGridView3.Rows[x].Cells[3].Value.ToString();

                }
            }
        }
        public void relleno_dctequebrado()
        {
            for (int i = 1; i < dataGridView14.ColumnCount; i++)
            {
                traer_estimacion(dataGridView14.Columns[i].Name, "DCTE QUEBRADO");
                for (int x = 0; x < dataGridView3.Rows.Count; x++)
                {
                    dataGridView14.Rows[x].Cells[i].Value = dataGridView3.Rows[x].Cells[3].Value.ToString();

                }
            }
        }
        public void  rellenar()
        {
            DataGridViewTextBoxColumn col12 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col13 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col14 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col15 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col16 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col17 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col18 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col19 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col20 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col21 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col22 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col23 = new DataGridViewTextBoxColumn();

            col12.HeaderText = "Productores";
            col13.HeaderText = "Productores";
            col14.HeaderText = "Productores";
            col15.HeaderText = "Productores";
            col16.HeaderText = "Productores";
            col17.HeaderText = "Productores";
            col18.HeaderText = "Productores";
            col19.HeaderText = "Productores";
            col20.HeaderText = "Productores";
            col21.HeaderText = "Productores";
            col22.HeaderText = "Productores";
            col23.HeaderText = "Productores";

            dataGridView2.Columns.Add(col12);
            dataGridView4.Columns.Add(col13);
            dataGridView5.Columns.Add(col14);
            dataGridView6.Columns.Add(col15);
            dataGridView7.Columns.Add(col16);
            dataGridView8.Columns.Add(col17);
            dataGridView9.Columns.Add(col18);
            dataGridView10.Columns.Add(col19);
            dataGridView11.Columns.Add(col20);
            dataGridView12.Columns.Add(col21);
            dataGridView13.Columns.Add(col22);
            dataGridView14.Columns.Add(col23);

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
           
                DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn col1 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn col2 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn col3 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn col4 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn col5 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn col6 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn col7 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn col8 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn col9 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn col10 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn col11 = new DataGridViewTextBoxColumn();
               
                col.HeaderText = dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                col.Name = dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                col1.HeaderText = dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                col1.Name = dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                col2.HeaderText = dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                col2.Name = dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                col3.HeaderText = dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                col3.Name = dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                col4.HeaderText = dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                col4.Name = dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                col5.HeaderText = dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                col5.Name = dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                col6.HeaderText = dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                col6.Name = dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                col7.HeaderText = dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                col7.Name= dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                col8.HeaderText = dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                col8.Name = dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                col9.HeaderText = dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                col9.Name = dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                col10.HeaderText = dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                col10.HeaderText = dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                col11.HeaderText = dataGridView1.Rows[i].Cells["productor"].Value.ToString();
                col11.Name = dataGridView1.Rows[i].Cells["productor"].Value.ToString();

                dataGridView2.Columns.Add(col);
                dataGridView4.Columns.Add(col1);
                dataGridView5.Columns.Add(col2);
                dataGridView6.Columns.Add(col3);
                dataGridView7.Columns.Add(col4);
                dataGridView8.Columns.Add(col5);
                dataGridView9.Columns.Add(col6);
                dataGridView10.Columns.Add(col7);
                dataGridView11.Columns.Add(col8);
                dataGridView12.Columns.Add(col9);
                dataGridView13.Columns.Add(col10);
                dataGridView14.Columns.Add(col11);
            }
      

        }

        private void button1_Click(object sender, EventArgs e)
        {
              traer_estimacion(dataGridView1.Rows[0].Cells[0].Value.ToString(), "SERR");
              rellenar_serr(dataGridView1.Rows[0].Cells[0].Value.ToString());
              relleno_Serr();
         
      
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("entro");
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(tabControl1.SelectedTab == tabControl1.TabPages[0])//your specific tabname
            {
                // MessageBox.Show("en");
                if (dataGridView2.Rows.Count == 0)
                {
                    traer_estimacion(dataGridView1.Rows[0].Cells[0].Value.ToString(), "SERR");
                    rellenar_serr(dataGridView1.Rows[0].Cells[0].Value.ToString());
                    relleno_Serr();
                }
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages[1])//your specific tabname
            {
                 // MessageBox.Show("en");
                 if(dataGridView4.Rows.Count == 0 )
                { 
                traer_estimacion(dataGridView1.Rows[0].Cells[0].Value.ToString(), "CHANDLER");
                rellenar_CHARLERD();
                relleno_chandler();}
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages[2])//your specific tabname
            {
                // MessageBox.Show("en");
                if(dataGridView5.Rows.Count == 0)
                { 
                 traer_estimacion(dataGridView1.Rows[0].Cells[0].Value.ToString(), "HOWARD");
                    rellenar_howard();
                    relleno_howard();
                }
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages[3])//your specific tabname
            {
                // MessageBox.Show("en");
                if (dataGridView6.Rows.Count == 0)
                {
                    traer_estimacion(dataGridView1.Rows[0].Cells[0].Value.ToString(), "TULARE");
                    rellenar_tulare();
                    relleno_tulare();
                }
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages[4])//your specific tabname
            {
                // MessageBox.Show("en");
                if (dataGridView7.Rows.Count == 0)
                {
                    traer_estimacion(dataGridView1.Rows[0].Cells[0].Value.ToString(), "SUNLAND");
                    rellenar_sundland();
                    relleno_sundland();
                }
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages[5])//your specific tabname
            {
                // MessageBox.Show("en");
                if (dataGridView8.Rows.Count == 0)
                {
                    traer_estimacion(dataGridView1.Rows[0].Cells[0].Value.ToString(), "HARTLEY");
                    rellenar_hartley();
                    relleno_hartley();
                }
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages[6])//your specific tabname
            {
                // MessageBox.Show("en");
                if (dataGridView9.Rows.Count == 0)
                {
                    traer_estimacion(dataGridView1.Rows[0].Cells[0].Value.ToString(), "SEMILLA");
                    rellenar_semilla();
                    relleno_semilla();
                }
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages[7])//your specific tabname
            {
                // MessageBox.Show("en");
                if (dataGridView10.Rows.Count == 0)
                {
                    traer_estimacion(dataGridView1.Rows[0].Cells[0].Value.ToString(), "VINA");
                    rellenar_vina();
                    relleno_vina();
                }
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages[8])//your specific tabname
            {
                // MessageBox.Show("en");
                if (dataGridView11.Rows.Count == 0)
                {
                    traer_estimacion(dataGridView1.Rows[0].Cells[0].Value.ToString(), "VINA MEJORADA");
                    rellenar_vina_mejorada();
                    relleno_vina_mejorada();
                }
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages[9])//your specific tabname
            {
                // MessageBox.Show("en");
                if (dataGridView12.Rows.Count == 0)
                {
                    traer_estimacion(dataGridView1.Rows[0].Cells[0].Value.ToString(), "FRANQUETTE");
                    rellenar_franquette();
                    relleno_franquette();
                }
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages[10])//your specific tabname
            {
                // MessageBox.Show("en");
                if (dataGridView13.Rows.Count == 0)
                {
                    traer_estimacion(dataGridView1.Rows[0].Cells[0].Value.ToString(), "DCTE DESPELONADO");
                    rellenar_dctedespelonado();
                    relleno_dctedespelonado();
                }
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages[11])//your specific tabname
            {
                // MessageBox.Show("en");
                if (dataGridView14.Rows.Count == 0)
                {
                    traer_estimacion(dataGridView1.Rows[0].Cells[0].Value.ToString(), "DCTE QUEBRADO");
                    rellenar_dctequebrado();
                    relleno_dctequebrado();
                }
            }

        }
    }
}
