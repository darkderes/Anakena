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
            this.estimacionTemporadaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.estimacionCalibreToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.estimacionCategoriasToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
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
            this.estimacionCategoriasToolStripMenuItem});
            this.estimacionesToolStripMenuItem.Name = "estimacionesToolStripMenuItem";
            this.estimacionesToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.E)));
            this.estimacionesToolStripMenuItem.Size = new System.Drawing.Size(88, 20);
            this.estimacionesToolStripMenuItem.Text = "Estimaciones";
            this.estimacionesToolStripMenuItem.Click += new System.EventHandler(this.estimacionesToolStripMenuItem_Click);
            // 
            // estimacionTemporadaToolStripMenuItem
            // 
            this.estimacionTemporadaToolStripMenuItem.Name = "estimacionTemporadaToolStripMenuItem";
            this.estimacionTemporadaToolStripMenuItem.Size = new System.Drawing.Size(297, 22);
            this.estimacionTemporadaToolStripMenuItem.Text = "E&stimacion Temporada                    ALT+S";
            this.estimacionTemporadaToolStripMenuItem.Click += new System.EventHandler(this.estimacionTemporadaToolStripMenuItem_Click);
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
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(585, 721);
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
            this.label3.Location = new System.Drawing.Point(656, 763);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(193, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Jorge Darderes 2015 Version 1.0";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(287, 209);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(125, 41);
            this.button1.TabIndex = 5;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // pictureBox2
            // 
            this.pictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox2.Image")));
            this.pictureBox2.Location = new System.Drawing.Point(659, 232);
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
            this.pictureBox1.Location = new System.Drawing.Point(620, 112);
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
            this.ClientSize = new System.Drawing.Size(1105, 812);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.pictureBox1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FormPrincipal";
            this.Text = "FormPrincipal";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
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
        private System.Windows.Forms.Button button1;
    }
}