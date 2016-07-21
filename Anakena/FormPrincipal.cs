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
            Cursor = Cursors.WaitCursor;
            FormCalibre_Estimacion s = new FormCalibre_Estimacion();
            s.ShowDialog();
            Cursor = Cursors.Default;
        }

        private void estimacionesToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void estimacionCategoriasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            FormCategoria_Estimacion s = new FormCategoria_Estimacion();
            s.ShowDialog();
            Cursor = Cursors.Default;
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

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void estimacionVSRealidadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormRealidadVSEstimacion s = new FormRealidadVSEstimacion();
            s.ShowDialog();
        }

        private void resumenEstimacionRealidadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormResumenEstimacionMasReal s = new FormResumenEstimacionMasReal();
            s.ShowDialog();
        }

        private void calibradoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormCalibrado s = new FormCalibrado();
            s.acceso = 0;
            s.ShowDialog();                
        }

        private void calibradorealEstimadoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormCalibrado s = new FormCalibrado();
            s.acceso = 1;
            s.ShowDialog();
        }

        private void kGCATCALToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormKGcat_cal s = new FormKGcat_cal();
            s.ShowDialog();
        }

        private void tarifasPorCalibreToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormTarifas_Calibre c = new FormTarifas_Calibre();
            c.ShowDialog();
        }

        private void tarifasPCCToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormProceso_CC c = new FormProceso_CC();
            c.ShowDialog();
        }

        private void tarifasPSCToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormProceso_SC c = new FormProceso_SC();
            c.ShowDialog();
        }

        private void preciosNCCToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormPrecioNCC s = new FormPrecioNCC();
            s.ShowDialog();
        }

        private void preciosSCCToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormPrecioNSC s = new FormPrecioNSC();
            s.ShowDialog();
        }

        private void prePorVarNCCToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormPor_Var v = new FormPor_Var();
            v.ShowDialog();
        }
    }
}
