namespace Anakena
{
    partial class FormRealidad
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormRealidad));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.BtnFiltro = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.cmb_variedad = new System.Windows.Forms.ComboBox();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.dataGridView3 = new System.Windows.Forms.DataGridView();
            this.dataGridView4 = new System.Windows.Forms.DataGridView();
            this.pBar1 = new System.Windows.Forms.ProgressBar();
            this.Btn_Excel = new System.Windows.Forms.Button();
            this.CmbTipo = new System.Windows.Forms.ComboBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.label2 = new System.Windows.Forms.Label();
            this.btnCancelar = new System.Windows.Forms.Button();
            this.dataGridView5 = new System.Windows.Forms.DataGridView();
            this.dataGridView6 = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView4)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView6)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.dataGridView1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(26, 116);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(1058, 478);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.Visible = false;
            // 
            // BtnFiltro
            // 
            this.BtnFiltro.Location = new System.Drawing.Point(747, 38);
            this.BtnFiltro.Name = "BtnFiltro";
            this.BtnFiltro.Size = new System.Drawing.Size(75, 23);
            this.BtnFiltro.TabIndex = 20;
            this.BtnFiltro.Text = "Buscar";
            this.BtnFiltro.UseVisualStyleBackColor = true;
            this.BtnFiltro.Click += new System.EventHandler(this.BtnFiltro_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Cursor = System.Windows.Forms.Cursors.No;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(374, 41);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(80, 16);
            this.label1.TabIndex = 19;
            this.label1.Text = "Variedad :";
            // 
            // cmb_variedad
            // 
            this.cmb_variedad.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmb_variedad.FormattingEnabled = true;
            this.cmb_variedad.Location = new System.Drawing.Point(460, 40);
            this.cmb_variedad.Name = "cmb_variedad";
            this.cmb_variedad.Size = new System.Drawing.Size(166, 21);
            this.cmb_variedad.TabIndex = 18;
            // 
            // dataGridView2
            // 
            this.dataGridView2.AllowUserToAddRows = false;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.dataGridView2.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Location = new System.Drawing.Point(26, 116);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.Size = new System.Drawing.Size(1079, 496);
            this.dataGridView2.TabIndex = 21;
            this.dataGridView2.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dataGridView2_CellFormatting);
            // 
            // dataGridView3
            // 
            this.dataGridView3.AllowUserToAddRows = false;
            this.dataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView3.Location = new System.Drawing.Point(42, 27);
            this.dataGridView3.Name = "dataGridView3";
            this.dataGridView3.Size = new System.Drawing.Size(214, 47);
            this.dataGridView3.TabIndex = 22;
            this.dataGridView3.Visible = false;
            // 
            // dataGridView4
            // 
            this.dataGridView4.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView4.Location = new System.Drawing.Point(844, 12);
            this.dataGridView4.Name = "dataGridView4";
            this.dataGridView4.Size = new System.Drawing.Size(240, 150);
            this.dataGridView4.TabIndex = 23;
            this.dataGridView4.Visible = false;
            this.dataGridView4.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView4_CellContentClick);
            // 
            // pBar1
            // 
            this.pBar1.Location = new System.Drawing.Point(879, 631);
            this.pBar1.Name = "pBar1";
            this.pBar1.Size = new System.Drawing.Size(226, 23);
            this.pBar1.TabIndex = 24;
            // 
            // Btn_Excel
            // 
            this.Btn_Excel.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("Btn_Excel.BackgroundImage")));
            this.Btn_Excel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.Btn_Excel.Location = new System.Drawing.Point(539, 631);
            this.Btn_Excel.Name = "Btn_Excel";
            this.Btn_Excel.Size = new System.Drawing.Size(47, 54);
            this.Btn_Excel.TabIndex = 26;
            this.Btn_Excel.UseVisualStyleBackColor = true;
            this.Btn_Excel.Visible = false;
            this.Btn_Excel.Click += new System.EventHandler(this.Btn_Excel_Click);
            // 
            // CmbTipo
            // 
            this.CmbTipo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CmbTipo.FormattingEnabled = true;
            this.CmbTipo.Items.AddRange(new object[] {
            "Kg",
            "%"});
            this.CmbTipo.Location = new System.Drawing.Point(641, 40);
            this.CmbTipo.Name = "CmbTipo";
            this.CmbTipo.Size = new System.Drawing.Size(90, 21);
            this.CmbTipo.TabIndex = 27;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.radioButton2);
            this.groupBox1.Controls.Add(this.radioButton1);
            this.groupBox1.Location = new System.Drawing.Point(377, 76);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(445, 34);
            this.groupBox1.TabIndex = 28;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Vista de reporte";
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.Location = new System.Drawing.Point(269, 11);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(71, 17);
            this.radioButton2.TabIndex = 1;
            this.radioButton2.TabStop = true;
            this.radioButton2.Text = "Agrupado";
            this.radioButton2.UseVisualStyleBackColor = true;
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Location = new System.Drawing.Point(122, 11);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(87, 17);
            this.radioButton1.TabIndex = 0;
            this.radioButton1.TabStop = true;
            this.radioButton1.Text = "Generalizado";
            this.radioButton1.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(876, 615);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(35, 13);
            this.label2.TabIndex = 25;
            this.label2.Text = "label2";
            // 
            // btnCancelar
            // 
            this.btnCancelar.Location = new System.Drawing.Point(952, 671);
            this.btnCancelar.Name = "btnCancelar";
            this.btnCancelar.Size = new System.Drawing.Size(75, 23);
            this.btnCancelar.TabIndex = 29;
            this.btnCancelar.Text = "cancelar";
            this.btnCancelar.UseVisualStyleBackColor = true;
            this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
            // 
            // dataGridView5
            // 
            this.dataGridView5.AllowUserToAddRows = false;
            dataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.dataGridView5.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle3;
            this.dataGridView5.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView5.Location = new System.Drawing.Point(26, 116);
            this.dataGridView5.Name = "dataGridView5";
            this.dataGridView5.Size = new System.Drawing.Size(1079, 196);
            this.dataGridView5.TabIndex = 30;
            this.dataGridView5.Visible = false;
            this.dataGridView5.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dataGridView5_CellFormatting);
            // 
            // dataGridView6
            // 
            this.dataGridView6.AllowUserToAddRows = false;
            dataGridViewCellStyle4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.dataGridView6.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle4;
            this.dataGridView6.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView6.Location = new System.Drawing.Point(26, 354);
            this.dataGridView6.Name = "dataGridView6";
            this.dataGridView6.Size = new System.Drawing.Size(1079, 196);
            this.dataGridView6.TabIndex = 31;
            this.dataGridView6.Visible = false;
            this.dataGridView6.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dataGridView6_CellFormatting);
            // 
            // FormRealidad
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.ClientSize = new System.Drawing.Size(1117, 706);
            this.Controls.Add(this.dataGridView6);
            this.Controls.Add(this.dataGridView5);
            this.Controls.Add(this.btnCancelar);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.CmbTipo);
            this.Controls.Add(this.Btn_Excel);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.pBar1);
            this.Controls.Add(this.dataGridView4);
            this.Controls.Add(this.dataGridView3);
            this.Controls.Add(this.dataGridView2);
            this.Controls.Add(this.BtnFiltro);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cmb_variedad);
            this.Controls.Add(this.dataGridView1);
            this.Name = "FormRealidad";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FormRealidad";
            this.Load += new System.EventHandler(this.FormRealidad_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView4)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView6)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button BtnFiltro;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cmb_variedad;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.DataGridView dataGridView3;
        private System.Windows.Forms.DataGridView dataGridView4;
        private System.Windows.Forms.ProgressBar pBar1;
        private System.Windows.Forms.Button Btn_Excel;
        private System.Windows.Forms.ComboBox CmbTipo;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton radioButton2;
        private System.Windows.Forms.RadioButton radioButton1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnCancelar;
        private System.Windows.Forms.DataGridView dataGridView5;
        private System.Windows.Forms.DataGridView dataGridView6;
    }
}