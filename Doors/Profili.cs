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
    public partial class Profili : Form
    {
        public Profili()
        {
            InitializeComponent();
        }
        public string ConnectionString = @"Data Source=.\SQLEXPRESS;AttachDbFilename=" + System.Windows.Forms.Application.StartupPath + @"\Resources\doors.mdf;Integrated Security=True;Connect Timeout=30";
        public void ShowData()
        {
            SqlConnection Connection = new SqlConnection(ConnectionString);
            try
            {
                Connection.Open();
            }
            catch (SqlException)
            {
                MessageBox.Show("Проверьте, достаточно ли места на диске, достаточно ли прав у учетной записи для операций с БД (См. справку), файлы MDF и LDF не должны быть помечены \"Только для чтения\". \n\nВозможно стоит попробовать отключить БД и запустить программу еще раз.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            SqlCommand comm = new SqlCommand("SELECT count(id_profil) FROM Profili", Connection);
            int kol = (int)comm.ExecuteScalar();
            if (kol > 0)
            {
                dataGridView1.RowCount = kol;
            }
            else
            {
                dataGridView1.RowCount = 1;
            }
            SqlCommand comm1 = new SqlCommand("SELECT * FROM Profili", Connection);
            SqlDataReader myReader = comm1.ExecuteReader(CommandBehavior.CloseConnection);
            int i = 0;
            while (myReader.Read())
            {
                for (int j = 0; j < 2; j++)
                {
                    dataGridView1[j, i].Value = Convert.ToString(myReader[j + 1]);
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
            string DelId = Convert.ToString(dataGridView1.CurrentCell.RowIndex + 1);
            string TextCommand = $"Delete from Profili where id_profil='{DelId}'";
            SqlCommand Command = new SqlCommand(TextCommand, Connection);
            try
            {
                Connection.Open();
            }
            catch (SqlException)
            {
                MessageBox.Show("Проверьте, достаточно ли места на диске, достаточно ли прав у учетной записи для операций с БД (См. справку), файлы MDF и LDF не должны быть помечены \"Только для чтения\". \n\nВозможно стоит попробовать отключить БД и запустить программу еще раз.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            Command.ExecuteNonQuery();
            Connection.Close();
            MessageBox.Show("Данные удалены");
            ShowData();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SqlConnection Connection = new SqlConnection(ConnectionString);
            try
            {
                Connection.Open();
            }
            catch (SqlException)
            {
                MessageBox.Show("Проверьте, достаточно ли места на диске, достаточно ли прав у учетной записи для операций с БД (См. справку), файлы MDF и LDF не должны быть помечены \"Только для чтения\". \n\nВозможно стоит попробовать отключить БД и запустить программу еще раз.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            SqlCommand comm = new SqlCommand("SELECT max(id_profil) FROM Profili", Connection);
            int max = (int)comm.ExecuteScalar();
            string sql = string.Format("Insert Into Profili" +
                  "(id_profil, profil,cena) Values(@kod, @pr,@c)");
            SqlCommand cmd = new SqlCommand(sql, Connection);
            cmd.Parameters.AddWithValue("@kod", Convert.ToString(max + 1));
            cmd.Parameters.AddWithValue("@pr", textBox1.Text);
            cmd.Parameters.AddWithValue("@c", textBox2.Text);
            if (cmd.ExecuteNonQuery() == 1)
            {
                MessageBox.Show("Запись успешно добавлена.");
            }
            ShowData();
            Connection.Close();
            textBox1.Clear();
            textBox2.Clear();
        }

        private void SomethingChanged(object sender, EventArgs e)
        {
            if (Is_NameOfProfil_Correct(textBox1.Text) && Is_CenaOfProfil_Correct(textBox2.Text))
            {
                button1.Visible = true;
            }
            else
            {
                button1.Visible = false;
            }
        }

        private bool Is_NameOfProfil_Correct(string text)
        {
            if (text.Length == 0 || string.IsNullOrWhiteSpace(text))
            {
                return false;
            }
            else
            {
                for (int i = 0; i < text.Length; i++)
                {
                    char symbol = text[i];
                    if (char.IsLetterOrDigit(symbol) || symbol == ' ' || symbol == '-' || symbol == '_')
                    {
                        if (i == text.Length - 1)
                        {
                            return true;
                        }
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            return false;
        }

        private bool Is_CenaOfProfil_Correct(string text)
        {
            if (text.Length == 0 || string.IsNullOrWhiteSpace(text))
            {
                return false;
            }
            else
            {
                for (int i = 0; i < text.Length; i++)
                {
                    char symbol = text[i];
                    if (char.IsDigit(symbol) || symbol == '.')
                    {
                        if (i == text.Length - 1)
                        {
                            string cena_corrupted = text.Replace('.', ',');
                            double cena = Convert.ToDouble(cena_corrupted + "0");
                            if (cena == 0)
                            {
                                return false;
                            }
                            else
                            {
                                return true;
                            }
                        }
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            return false;
        }

        private void Profili_HelpRequested(object sender, HelpEventArgs hlpevent)
        {
            System.Diagnostics.Process.Start(Environment.CurrentDirectory + "\\Resources\\HelpFile.chm");
        }
    }
}
