﻿using System;
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

        private void button2_Click(object sender, EventArgs e)
        {
            Dop f1 = new Dop();
            f1.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Add_zakaz f1 = new Add_zakaz();
            f1.Show();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Calc f1 = new Calc();
            f1.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(Environment.CurrentDirectory + "\\Resources\\HelpFile.chm");
        }

        private void Main_HelpRequested(object sender, HelpEventArgs hlpevent)
        {
            System.Diagnostics.Process.Start(Environment.CurrentDirectory + "\\Resources\\HelpFile.chm");
        }
    }
}
