using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Doors
{
    public partial class Add_zakaz : Form
    {
        public Add_zakaz()
        {
            InitializeComponent();
        }
        public string ConnectionString = @"Data Source=.\SQLEXPRESS;AttachDbFilename=" + Application.StartupPath + @"\doors.mdf;Integrated Security=True;Connect Timeout=30";
        public void Selpr()
        {
            SqlConnection Connection = new SqlConnection(ConnectionString);

            Connection.Open();

            SqlCommand comm1 = new SqlCommand("SELECT profil FROM Profili", Connection);
            SqlDataReader myReader = comm1.ExecuteReader(CommandBehavior.CloseConnection);

            comboBox1.Items.Clear();

            while (myReader.Read())
            {
                comboBox1.Items.Add(myReader[0]);
            }

            comboBox1.SelectedIndex = 0;

            Connection.Close();
        }
        private void Add_zakaz_Load(object sender, EventArgs e)
        {
            Selpr();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SqlConnection Connection = new SqlConnection(ConnectionString);

            Connection.Open();
            SqlCommand comm = new SqlCommand("SELECT max(id_zakaz) FROM Zakazy", Connection);

            int max = (int)comm.ExecuteScalar();
            string sql = string.Format("Insert Into Zakazy" +
                  "(id_zakaz, data, id_profil, vysota, shirina, kolvo,ustanovka,nalichniki,zamok,ruchka,petli) Values(@kod, @data, @pr, @v, @sh, @k, @u, @nal, @z, @r, @p)");
            SqlCommand cmd = new SqlCommand(sql, Connection);
            cmd.Parameters.AddWithValue("@kod", Convert.ToString(max + 1));
            cmd.Parameters.AddWithValue("@data", dateTimePicker1.Value);
            cmd.Parameters.AddWithValue("@pr", Convert.ToString(comboBox1.SelectedIndex + 1));
            cmd.Parameters.AddWithValue("@v", comboBox2.Text);
            cmd.Parameters.AddWithValue("@sh", comboBox3.Text);
            cmd.Parameters.AddWithValue("@k", numericUpDown1.Text);
            cmd.Parameters.AddWithValue("@u", checkBox1.Checked);
            cmd.Parameters.AddWithValue("@nal", checkBox2.Checked);
            cmd.Parameters.AddWithValue("@z", checkBox3.Checked);
            cmd.Parameters.AddWithValue("@r", checkBox4.Checked);
            cmd.Parameters.AddWithValue("@p", checkBox5.Checked);
            if (cmd.ExecuteNonQuery() == 1)
            {
                MessageBox.Show("Запись успешно добавлена.");
            }

            Connection.Close();
        }
    }
}

