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
    public partial class Profili : Form
    {
        public Profili()
        {
            InitializeComponent();
        }
        public string ConnectionString = @"Data Source=.\SQLEXPRESS;AttachDbFilename=" + System.Windows.Forms.Application.StartupPath + @"\DOORS.mdf;Integrated Security=True;Connect Timeout=30";
        public void ShowData()
        {


            SqlConnection Connection = new SqlConnection(ConnectionString);

            Connection.Open();

            SqlCommand comm = new SqlCommand("SELECT count(id_profil) FROM Profili", Connection);
            Int32 kol = (Int32)comm.ExecuteScalar();
            if (kol > 0)
                dataGridView1.RowCount = kol;
            else dataGridView1.RowCount = 1;

            SqlCommand comm1 = new SqlCommand("SELECT * FROM Profili", Connection);
            SqlDataReader myReader = comm1.ExecuteReader(CommandBehavior.CloseConnection);

            int i = 0;
            while (myReader.Read())
            {
                for (int j = 0; j < 3; j++)
                {

                    dataGridView1[j, i].Value = Convert.ToString(myReader[j]);

                }
                i++;

            }
            Connection.Close();
        }
        private void Profili_Load(object sender, EventArgs e)
        {
            ShowData();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SqlConnection Connection = new SqlConnection(ConnectionString);

            string DelId = Convert.ToString(dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value);


            string TextCommand = "Delete from Profili where id_profil=" + DelId;

            SqlCommand Command = new SqlCommand(TextCommand, Connection);

            Connection.Open();

            Command.ExecuteNonQuery();

            Connection.Close();

            MessageBox.Show("данные удалены");

            ShowData();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SqlConnection Connection = new SqlConnection(ConnectionString);

            Connection.Open();
            SqlCommand comm = new SqlCommand("SELECT max(id_profil) FROM Profili", Connection);

            Int32 max = (Int32)comm.ExecuteScalar();
            string sql = string.Format("Insert Into Profili" +
                  "(id_profil, profil,cena) Values(@kod, @pr,@c)");
            SqlCommand cmd = new SqlCommand(sql, Connection);
            cmd.Parameters.AddWithValue("@kod", Convert.ToString(max + 1));
            cmd.Parameters.AddWithValue("@pr", textBox1.Text);
            cmd.Parameters.AddWithValue("@c", textBox2.Text);
            if (cmd.ExecuteNonQuery() == 1)
                MessageBox.Show("Запись успешно добавлена.");
            ShowData();
            Connection.Close();
            textBox1.Clear();
            textBox2.Clear();
        }
    }
}
