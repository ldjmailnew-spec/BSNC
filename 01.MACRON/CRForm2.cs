using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
//using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CRPRT
{
    public partial class CRForm2 : Form
    {
        [STAThread]
        static void Main()
        {
        }

        public CRForm2()
        {
            InitializeComponent();
            CrView2.Zoom(100);
        }

        private void CrView2_Load(object sender, EventArgs e)
        {

        }

    }
}

