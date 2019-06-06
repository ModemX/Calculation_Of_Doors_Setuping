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
        public string ConnectionString = @"Data Source=.\SQLEXPRESS;AttachDbFilename=" + System.Windows.Forms.Application.StartupPath + @"\Resources\doors.mdf;Integrated Security=True;Connect Timeout=30";
        public void Selpr()
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
            try
            {
                Connection.Open();
            }
            catch (SqlException)
            {
                MessageBox.Show("Проверьте, достаточно ли места на диске, достаточно ли прав у учетной записи для операций с БД (См. справку), файлы MDF и LDF не должны быть помечены \"Только для чтения\". \n\nВозможно стоит попробовать отключить БД и запустить программу еще раз.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            SqlCommand comm = new SqlCommand("SELECT max(id_zakaz) FROM Zakazy", Connection);
            int max;
            object temp = comm.ExecuteScalar();
            if (temp is DBNull)
            {
                max = 0;
            }
            else
            {
                max = (int)temp;
            }

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
                NewBlank blank = new NewBlank(max, dateTimePicker1.Value, Convert.ToInt32(comboBox1.SelectedIndex), Convert.ToInt32(comboBox2.Text), Convert.ToInt32(comboBox3.Text), Convert.ToInt32(numericUpDown1.Text), checkBox1.Checked, checkBox2.Checked, checkBox3.Checked, checkBox4.Checked, checkBox5.Checked, Заказчик_ФИО.Text, Convert.ToInt32(Заказчик_ТелефонныйПрефикс.Text), Convert.ToInt32(Заказчик_Телефон.Text));
                MessageBox.Show("Запись успешно добавлена.");
            }
            Connection.Close();
        }

        private void CheckIsEverythingCorrect()
        {
            if (numericUpDown1.Value > 0 &&
                IsStringCorrectWordsOnly(comboBox1.Text) &&
                IsStringCorrectPositiveDigitsOnly(comboBox2.Text) &&
                IsStringCorrectPositiveDigitsOnly(comboBox3.Text) &&
                IsStringCorrectWordsOnly(Заказчик_ФИО.Text) &&
                IsStringCorrectPhonePrefix(Заказчик_ТелефонныйПрефикс.Text) &&
                IsStringCorrectPhone(Заказчик_Телефон.Text))
            {
                button1.Enabled = true;
            }
            else
            {
                button1.Enabled = false;
            }
        }

        private bool IsStringCorrectPhone(string text)
        {
            if (text.Length == 7)
            {
                for (int i = 0; i < text.Length; i++)
                {
                    char symbol = text[i];
                    if (char.IsNumber(symbol))
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

        private bool IsStringCorrectPhonePrefix(string text)
        {
            if (text == "17" || text == "29" || text == "33" || text == "44")
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private void SomethingChanged(object sender, EventArgs e)
        {
            CheckIsEverythingCorrect();
        }

        private bool IsStringCorrectWordsOnly(string text)
        {
            for (int i = 0; i < text.Length; i++)
            {
                char symbol = text[i];
                if (char.IsLetter(symbol) || symbol == ' ')
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
            return false;
        }

        private bool IsStringCorrectPositiveDigitsOnly(string text)
        {
            for (int i = 0; i < text.Length; i++)
            {
                char symbol = text[i];
                if (char.IsNumber(symbol))
                {
                    if (i < text.Length)
                    {
                        int number = Convert.ToInt32(text);
                        if (number >= 50)
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
            return false;
        }

        private void Add_zakaz_HelpRequested(object sender, HelpEventArgs hlpevent)
        {
            System.Diagnostics.Process.Start(Environment.CurrentDirectory + "\\Resources\\HelpFile.chm");
        }
    }
}