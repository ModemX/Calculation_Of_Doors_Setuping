using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.Sql;
using System.Data.SqlClient;

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

        private void domainUpDown1_SelectedItemChanged(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

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

            Int32 max = (Int32)comm.ExecuteScalar();
            string sql = string.Format("Insert Into Zakazy" +
                  "(id_zakaz, data, id_profil, vysota, shirina, kolvo,ustanovka,nalichniki,zamok,ruchka,petli) Values(@kod, @data, @pr, @v, @sh, @k, @u, @nal, @z, @r, @p)");
            SqlCommand cmd = new SqlCommand(sql, Connection);
            cmd.Parameters.Add("@kod", Convert.ToString(max + 1));
            cmd.Parameters.Add("@data", dateTimePicker1.Value);
            cmd.Parameters.Add("@pr", Convert.ToString(comboBox1.SelectedIndex + 1));
            cmd.Parameters.Add("@v", comboBox2.Text);
            cmd.Parameters.Add("@sh", comboBox3.Text);
            cmd.Parameters.Add("@k", numericUpDown1.Text);
            cmd.Parameters.Add("@u", checkBox1.Checked);
            cmd.Parameters.Add("@nal", checkBox2.Checked);
            cmd.Parameters.Add("@z", checkBox3.Checked);
            cmd.Parameters.Add("@r", checkBox4.Checked);
            cmd.Parameters.Add("@p", checkBox5.Checked);
            if (cmd.ExecuteNonQuery() == 1)
                MessageBox.Show("Запись успешно добавлена.");
            Connection.Close();
        }
        }
    }

