using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VeriTabanıV1._3
{
    public partial class Form2akademi : Form
    {
        public Form2akademi()
        {
            InitializeComponent();
        }

        private void Form2akademi_Load(object sender, EventArgs e)
        {
            Form1 f1 = new Form1();
            f1.Close();
        }
    }
}
