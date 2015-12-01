using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Anakena
{
    public partial class FormPrincipal : Form
    {
        public FormPrincipal()
        {
            InitializeComponent();
        }

        private void estimacionTemporadaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form1 s = new Form1();
            s.ShowDialog();
        }

        private void estimacionCalibreToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormCalibre_Estimacion s = new FormCalibre_Estimacion();
            s.ShowDialog();
        }

        private void estimacionesToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void estimacionCategoriasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormCategoria_Estimacion s = new FormCategoria_Estimacion();
            s.ShowDialog();
        }
    }
}
