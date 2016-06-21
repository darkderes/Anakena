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
    public partial class FormLoading : Form
    {
        public FormLoading()
        {
            InitializeComponent();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void FormLoading_Load(object sender, EventArgs e)
        {
            pictureBox1.Image = Image.FromFile("C:/spin.gif");
        }
    }
}
