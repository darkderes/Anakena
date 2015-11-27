namespace Anakena
{
    partial class Form1
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.TxtFile = new System.Windows.Forms.Label();
            this.Btn_Examinar = new System.Windows.Forms.Button();
            this.CmbHojas = new System.Windows.Forms.ComboBox();
            this.Btn_Cargar = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.btn_ingresar = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader;
            this.dataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 92);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(1045, 547);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.CellBeginEdit += new System.Windows.Forms.DataGridViewCellCancelEventHandler(this.dataGridView1_CellBeginEdit);
            this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellClick);
            this.dataGridView1.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dataGridView1_CellFormatting);
            // 
            // TxtFile
            // 
            this.TxtFile.AutoSize = true;
            this.TxtFile.Location = new System.Drawing.Point(407, 30);
            this.TxtFile.Name = "TxtFile";
            this.TxtFile.Size = new System.Drawing.Size(35, 13);
            this.TxtFile.TabIndex = 1;
            this.TxtFile.Text = "label1";
            this.TxtFile.Visible = false;
            // 
            // Btn_Examinar
            // 
            this.Btn_Examinar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.Btn_Examinar.ImageAlign = System.Drawing.ContentAlignment.TopLeft;
            this.Btn_Examinar.Location = new System.Drawing.Point(436, 23);
            this.Btn_Examinar.Name = "Btn_Examinar";
            this.Btn_Examinar.Size = new System.Drawing.Size(186, 26);
            this.Btn_Examinar.TabIndex = 2;
            this.Btn_Examinar.Text = "Examinar";
            this.Btn_Examinar.UseVisualStyleBackColor = true;
            this.Btn_Examinar.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // CmbHojas
            // 
            this.CmbHojas.FormattingEnabled = true;
            this.CmbHojas.Location = new System.Drawing.Point(410, 27);
            this.CmbHojas.Name = "CmbHojas";
            this.CmbHojas.Size = new System.Drawing.Size(158, 21);
            this.CmbHojas.TabIndex = 3;
            this.CmbHojas.Visible = false;
            // 
            // Btn_Cargar
            // 
            this.Btn_Cargar.Location = new System.Drawing.Point(591, 23);
            this.Btn_Cargar.Name = "Btn_Cargar";
            this.Btn_Cargar.Size = new System.Drawing.Size(64, 26);
            this.Btn_Cargar.TabIndex = 4;
            this.Btn_Cargar.Text = "Cargar";
            this.Btn_Cargar.UseVisualStyleBackColor = true;
            this.Btn_Cargar.Visible = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(12, 12);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(173, 73);
            this.pictureBox1.TabIndex = 5;
            this.pictureBox1.TabStop = false;
            // 
            // btn_ingresar
            // 
            this.btn_ingresar.Location = new System.Drawing.Point(464, 646);
            this.btn_ingresar.Name = "btn_ingresar";
            this.btn_ingresar.Size = new System.Drawing.Size(130, 56);
            this.btn_ingresar.TabIndex = 6;
            this.btn_ingresar.Text = "INGRESAR";
            this.btn_ingresar.UseVisualStyleBackColor = true;
            this.btn_ingresar.Click += new System.EventHandler(this.btn_ingresar_Click);
            // 
            // Form1
            // 
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.ClientSize = new System.Drawing.Size(1069, 714);
            this.Controls.Add(this.btn_ingresar);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.Btn_Cargar);
            this.Controls.Add(this.CmbHojas);
            this.Controls.Add(this.Btn_Examinar);
            this.Controls.Add(this.TxtFile);
            this.Controls.Add(this.dataGridView1);
            this.ImeMode = System.Windows.Forms.ImeMode.On;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.Form1_Load_1);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }


        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label TxtFile;
        private System.Windows.Forms.Button Btn_Examinar;
        private System.Windows.Forms.ComboBox CmbHojas;
        private System.Windows.Forms.Button Btn_Cargar;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button btn_ingresar;
    }
}

