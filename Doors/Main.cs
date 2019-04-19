using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Doors
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Profili f1 = new Profili();
            f1.Show();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Uchet_zakaz f1 = new Uchet_zakaz();
            f1.Show();
        }
    }
}
