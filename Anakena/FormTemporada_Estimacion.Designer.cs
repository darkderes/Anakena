namespace Anakena
{
    partial class FormTemporada_Estimacion
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormTemporada_Estimacion));
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.TxtFile = new System.Windows.Forms.Label();
            this.CmbHojas = new System.Windows.Forms.ComboBox();
            this.Btn_Cargar = new System.Windows.Forms.Button();
            this.btn_ingresar = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.Btn_Examinar = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.dataGridView1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 92);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.Size = new System.Drawing.Size(1120, 547);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dataGridView1_CellFormatting);
            // 
            // TxtFile
            // 
            this.TxtFile.AutoSize = true;
            this.TxtFile.Location = new System.Drawing.Point(812, 43);
            this.TxtFile.Name = "TxtFile";
            this.TxtFile.Size = new System.Drawing.Size(35, 13);
            this.TxtFile.TabIndex = 1;
            this.TxtFile.Text = "label1";
            this.TxtFile.Visible = false;
            // 
            // CmbHojas
            // 
            this.CmbHojas.FormattingEnabled = true;
            this.CmbHojas.Location = new System.Drawing.Point(718, 12);
            this.CmbHojas.Name = "CmbHojas";
            this.CmbHojas.Size = new System.Drawing.Size(158, 21);
            this.CmbHojas.TabIndex = 3;
            this.CmbHojas.Visible = false;
            // 
            // Btn_Cargar
            // 
            this.Btn_Cargar.Location = new System.Drawing.Point(718, 30);
            this.Btn_Cargar.Name = "Btn_Cargar";
            this.Btn_Cargar.Size = new System.Drawing.Size(64, 26);
            this.Btn_Cargar.TabIndex = 4;
            this.Btn_Cargar.Text = "Cargar";
            this.Btn_Cargar.UseVisualStyleBackColor = true;
            this.Btn_Cargar.Visible = false;
            // 
            // btn_ingresar
            // 
            this.btn_ingresar.BackColor = System.Drawing.SystemColors.Control;
            this.btn_ingresar.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_ingresar.Image = global::Anakena.Properties.Resources.imagen_add;
            this.btn_ingresar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btn_ingresar.Location = new System.Drawing.Point(461, 664);
            this.btn_ingresar.Name = "btn_ingresar";
            this.btn_ingresar.Size = new System.Drawing.Size(209, 38);
            this.btn_ingresar.TabIndex = 6;
            this.btn_ingresar.Text = "Ingreso estimacion";
            this.btn_ingresar.UseVisualStyleBackColor = false;
            this.btn_ingresar.Visible = false;
            this.btn_ingresar.Click += new System.EventHandler(this.btn_ingresar_Click);
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
            // Btn_Examinar
            // 
            this.Btn_Examinar.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.Btn_Examinar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.Btn_Examinar.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Btn_Examinar.Image = global::Anakena.Properties.Resources.examinar_mod;
            this.Btn_Examinar.ImageAlign = System.Drawing.ContentAlignment.TopLeft;
            this.Btn_Examinar.Location = new System.Drawing.Point(484, 21);
            this.Btn_Examinar.Name = "Btn_Examinar";
            this.Btn_Examinar.Size = new System.Drawing.Size(186, 35);
            this.Btn_Examinar.TabIndex = 2;
            this.Btn_Examinar.Text = "Carga .XLS";
            this.Btn_Examinar.UseVisualStyleBackColor = false;
            this.Btn_Examinar.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // FormTemporada_Estimacion
            // 
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.ClientSize = new System.Drawing.Size(1144, 714);
            this.Controls.Add(this.btn_ingresar);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.Btn_Cargar);
            this.Controls.Add(this.CmbHojas);
            this.Controls.Add(this.Btn_Examinar);
            this.Controls.Add(this.TxtFile);
            this.Controls.Add(this.dataGridView1);
            this.ImeMode = System.Windows.Forms.ImeMode.On;
            this.Name = "FormTemporada_Estimacion";
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

