namespace Anakena
{
    partial class FormCalibrado
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormCalibrado));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.pictureBox3 = new System.Windows.Forms.PictureBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.cmb_variedad = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.BtnFiltro = new System.Windows.Forms.Button();
            this.Btn_Excel = new System.Windows.Forms.Button();
            this.pBar1 = new System.Windows.Forms.ProgressBar();
            this.dataGridView4 = new System.Windows.Forms.DataGridView();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBox2
            // 
            this.pictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox2.Image")));
            this.pictureBox2.Location = new System.Drawing.Point(348, 205);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(111, 81);
            this.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox2.TabIndex = 48;
            this.pictureBox2.TabStop = false;
            // 
            // pictureBox3
            // 
            this.pictureBox3.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBox3.BackgroundImage")));
            this.pictureBox3.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBox3.Location = new System.Drawing.Point(308, 116);
            this.pictureBox3.Name = "pictureBox3";
            this.pictureBox3.Size = new System.Drawing.Size(218, 76);
            this.pictureBox3.TabIndex = 47;
            this.pictureBox3.TabStop = false;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.cmb_variedad);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.BtnFiltro);
            this.groupBox1.Location = new System.Drawing.Point(258, 18);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(374, 51);
            this.groupBox1.TabIndex = 46;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Opciones de busqueda";
            // 
            // cmb_variedad
            // 
            this.cmb_variedad.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmb_variedad.FormattingEnabled = true;
            this.cmb_variedad.Location = new System.Drawing.Point(102, 19);
            this.cmb_variedad.Name = "cmb_variedad";
            this.cmb_variedad.Size = new System.Drawing.Size(166, 21);
            this.cmb_variedad.TabIndex = 27;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Cursor = System.Windows.Forms.Cursors.No;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(16, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(80, 16);
            this.label1.TabIndex = 28;
            this.label1.Text = "Variedad :";
            // 
            // BtnFiltro
            // 
            this.BtnFiltro.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.BtnFiltro.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnFiltro.Location = new System.Drawing.Point(284, 17);
            this.BtnFiltro.Name = "BtnFiltro";
            this.BtnFiltro.Size = new System.Drawing.Size(75, 23);
            this.BtnFiltro.TabIndex = 29;
            this.BtnFiltro.Text = "Buscar";
            this.BtnFiltro.UseVisualStyleBackColor = false;
            this.BtnFiltro.Click += new System.EventHandler(this.BtnFiltro_Click);
            // 
            // Btn_Excel
            // 
            this.Btn_Excel.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("Btn_Excel.BackgroundImage")));
            this.Btn_Excel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.Btn_Excel.Location = new System.Drawing.Point(394, 349);
            this.Btn_Excel.Name = "Btn_Excel";
            this.Btn_Excel.Size = new System.Drawing.Size(50, 51);
            this.Btn_Excel.TabIndex = 45;
            this.Btn_Excel.UseVisualStyleBackColor = true;
            this.Btn_Excel.Click += new System.EventHandler(this.Btn_Excel_Click);
            // 
            // pBar1
            // 
            this.pBar1.Location = new System.Drawing.Point(638, 361);
            this.pBar1.Name = "pBar1";
            this.pBar1.Size = new System.Drawing.Size(138, 10);
            this.pBar1.TabIndex = 44;
            this.pBar1.Visible = false;
            // 
            // dataGridView4
            // 
            this.dataGridView4.AllowUserToAddRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.dataGridView4.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView4.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridView4.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dataGridView4.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView4.Location = new System.Drawing.Point(59, 85);
            this.dataGridView4.Name = "dataGridView4";
            this.dataGridView4.Size = new System.Drawing.Size(717, 252);
            this.dataGridView4.TabIndex = 43;
            this.dataGridView4.Visible = false;
            this.dataGridView4.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dataGridView4_CellFormatting);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(59, 18);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(133, 47);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 42;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Visible = false;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.dataGridView1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(59, 91);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(717, 252);
            this.dataGridView1.TabIndex = 49;
            this.dataGridView1.Visible = false;
            // 
            // FormCalibrado
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.ClientSize = new System.Drawing.Size(834, 412);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.pictureBox3);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.Btn_Excel);
            this.Controls.Add(this.pBar1);
            this.Controls.Add(this.dataGridView4);
            this.Controls.Add(this.pictureBox1);
            this.Name = "FormCalibrado";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FormCalibrado";
            this.Load += new System.EventHandler(this.FormCalibrado_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.PictureBox pictureBox3;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ComboBox cmb_variedad;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button BtnFiltro;
        private System.Windows.Forms.Button Btn_Excel;
        private System.Windows.Forms.ProgressBar pBar1;
        private System.Windows.Forms.DataGridView dataGridView4;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.DataGridView dataGridView1;
    }
}