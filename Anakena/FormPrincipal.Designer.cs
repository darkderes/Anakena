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
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.estimacionesToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(830, 24);
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
            // pictureBox1
            // 
            this.pictureBox1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBox1.BackgroundImage")));
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBox1.Location = new System.Drawing.Point(620, 314);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(247, 119);
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            // 
            // FormPrincipal
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.ClientSize = new System.Drawing.Size(830, 491);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.pictureBox1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FormPrincipal";
            this.Text = "FormPrincipal";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
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
    }
}