namespace Anakena
{
    partial class FormPrincipal
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormPrincipal));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.estimacionesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.estimacionCalibreToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.estimacionCategoriasToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.estimacionCalibreCategoriaRecepcionToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.estimacionCalibreCategoriaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.estimacionTemporadaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.realToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.factorProductorVariedadToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.estimacionVSRealidadToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.salirToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.estimacionesToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1105, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // estimacionesToolStripMenuItem
            // 
            this.estimacionesToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.estimacionTemporadaToolStripMenuItem,
            this.estimacionCalibreToolStripMenuItem,
            this.estimacionCategoriasToolStripMenuItem,
            this.estimacionCalibreCategoriaRecepcionToolStripMenuItem,
            this.estimacionCalibreCategoriaToolStripMenuItem,
            this.realToolStripMenuItem,
            this.factorProductorVariedadToolStripMenuItem,
            this.estimacionVSRealidadToolStripMenuItem,
            this.salirToolStripMenuItem});
            this.estimacionesToolStripMenuItem.Name = "estimacionesToolStripMenuItem";
            this.estimacionesToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.E)));
            this.estimacionesToolStripMenuItem.Size = new System.Drawing.Size(88, 20);
            this.estimacionesToolStripMenuItem.Text = "Estimaciones";
            this.estimacionesToolStripMenuItem.Click += new System.EventHandler(this.estimacionesToolStripMenuItem_Click);
            // 
            // estimacionCalibreToolStripMenuItem
            // 
            this.estimacionCalibreToolStripMenuItem.Name = "estimacionCalibreToolStripMenuItem";
            this.estimacionCalibreToolStripMenuItem.Size = new System.Drawing.Size(297, 22);
            this.estimacionCalibreToolStripMenuItem.Text = "Es&timacion Calibre                            ALT+T";
            this.estimacionCalibreToolStripMenuItem.Click += new System.EventHandler(this.estimacionCalibreToolStripMenuItem_Click);
            // 
            // estimacionCategoriasToolStripMenuItem
            // 
            this.estimacionCategoriasToolStripMenuItem.Name = "estimacionCategoriasToolStripMenuItem";
            this.estimacionCategoriasToolStripMenuItem.Size = new System.Drawing.Size(297, 22);
            this.estimacionCategoriasToolStripMenuItem.Text = "Esti&macion Categorias                      ALT+M";
            this.estimacionCategoriasToolStripMenuItem.Click += new System.EventHandler(this.estimacionCategoriasToolStripMenuItem_Click);
            // 
            // estimacionCalibreCategoriaRecepcionToolStripMenuItem
            // 
            this.estimacionCalibreCategoriaRecepcionToolStripMenuItem.Name = "estimacionCalibreCategoriaRecepcionToolStripMenuItem";
            this.estimacionCalibreCategoriaRecepcionToolStripMenuItem.Size = new System.Drawing.Size(297, 22);
            this.estimacionCalibreCategoriaRecepcionToolStripMenuItem.Text = "Estimacion Calibre-Categoria-Recepcion";
            this.estimacionCalibreCategoriaRecepcionToolStripMenuItem.Click += new System.EventHandler(this.estimacionCalibreCategoriaRecepcionToolStripMenuItem_Click);
            // 
            // estimacionCalibreCategoriaToolStripMenuItem
            // 
            this.estimacionCalibreCategoriaToolStripMenuItem.Name = "estimacionCalibreCategoriaToolStripMenuItem";
            this.estimacionCalibreCategoriaToolStripMenuItem.Size = new System.Drawing.Size(297, 22);
            this.estimacionCalibreCategoriaToolStripMenuItem.Text = "Resumen Variedad";
            this.estimacionCalibreCategoriaToolStripMenuItem.Click += new System.EventHandler(this.estimacionCalibreCategoriaToolStripMenuItem_Click);
            // 
            // estimacionTemporadaToolStripMenuItem
            // 
            this.estimacionTemporadaToolStripMenuItem.Name = "estimacionTemporadaToolStripMenuItem";
            this.estimacionTemporadaToolStripMenuItem.Size = new System.Drawing.Size(297, 22);
            this.estimacionTemporadaToolStripMenuItem.Text = "Estimación Kg";
            this.estimacionTemporadaToolStripMenuItem.Click += new System.EventHandler(this.estimacionTemporadaToolStripMenuItem_Click);
            // 
            // realToolStripMenuItem
            // 
            this.realToolStripMenuItem.Name = "realToolStripMenuItem";
            this.realToolStripMenuItem.Size = new System.Drawing.Size(297, 22);
            this.realToolStripMenuItem.Text = "Real";
            this.realToolStripMenuItem.Click += new System.EventHandler(this.realToolStripMenuItem_Click);
            // 
            // factorProductorVariedadToolStripMenuItem
            // 
            this.factorProductorVariedadToolStripMenuItem.Name = "factorProductorVariedadToolStripMenuItem";
            this.factorProductorVariedadToolStripMenuItem.Size = new System.Drawing.Size(297, 22);
            this.factorProductorVariedadToolStripMenuItem.Text = "Factor Productor - Variedad";
            this.factorProductorVariedadToolStripMenuItem.Click += new System.EventHandler(this.factorProductorVariedadToolStripMenuItem_Click);
            // 
            // estimacionVSRealidadToolStripMenuItem
            // 
            this.estimacionVSRealidadToolStripMenuItem.Name = "estimacionVSRealidadToolStripMenuItem";
            this.estimacionVSRealidadToolStripMenuItem.Size = new System.Drawing.Size(297, 22);
            this.estimacionVSRealidadToolStripMenuItem.Text = "Estimacion V/S Realidad";
            this.estimacionVSRealidadToolStripMenuItem.Click += new System.EventHandler(this.estimacionVSRealidadToolStripMenuItem_Click);
            // 
            // salirToolStripMenuItem
            // 
            this.salirToolStripMenuItem.Name = "salirToolStripMenuItem";
            this.salirToolStripMenuItem.Size = new System.Drawing.Size(297, 22);
            this.salirToolStripMenuItem.Text = "Salir";
            this.salirToolStripMenuItem.Click += new System.EventHandler(this.salirToolStripMenuItem_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(475, 610);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(352, 16);
            this.label2.TabIndex = 1;
            this.label2.Text = "Desarrollado para Exportadora Anakena Limitada";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.Black;
            this.label3.Location = new System.Drawing.Point(546, 652);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(193, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Jorge Darderes 2015 Version 1.0";
            // 
            // pictureBox2
            // 
            this.pictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox2.Image")));
            this.pictureBox2.Location = new System.Drawing.Point(565, 244);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(160, 127);
            this.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox2.TabIndex = 2;
            this.pictureBox2.TabStop = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBox1.BackgroundImage")));
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBox1.Location = new System.Drawing.Point(516, 111);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(239, 92);
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            // 
            // FormPrincipal
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.ClientSize = new System.Drawing.Size(1105, 781);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.pictureBox1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FormPrincipal";
            this.Text = "Menu Principal";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.FormPrincipal_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem estimacionesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem estimacionTemporadaToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem estimacionCalibreToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem estimacionCategoriasToolStripMenuItem;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ToolStripMenuItem estimacionCalibreCategoriaRecepcionToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem estimacionCalibreCategoriaToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem realToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem factorProductorVariedadToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem salirToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem estimacionVSRealidadToolStripMenuItem;
    }
}