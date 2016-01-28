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
            FormTemporada_Estimacion s = new FormTemporada_Estimacion();
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

        private void button1_Click(object sender, EventArgs e)
        {
            FormEstimacion f = new FormEstimacion();
            f.ShowDialog();
        }

        private void estimacionCalibreCategoriaRecepcionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormEstimacion s = new FormEstimacion();
            s.ShowDialog();
        }

        private void FormPrincipal_Load(object sender, EventArgs e)
        {

        }

        private void estimacionCalibreCategoriaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormCC_Estimacion s = new FormCC_Estimacion();
            s.ShowDialog();
        }

        private void realToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormRealidad s = new FormRealidad();
            s.ShowDialog();
        }

        private void factorProductorVariedadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormFactor s = new FormFactor();
            s.ShowDialog();
        }
    }
}
