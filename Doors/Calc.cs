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
    public partial class Calc : Form
    {
        public Calc()
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
        public void SelCena()
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
            SqlCommand comm1 = new SqlCommand("SELECT cena FROM Profili", Connection);
            SqlDataReader myReader = comm1.ExecuteReader(CommandBehavior.CloseConnection);
            comboBox5.Items.Clear();
            while (myReader.Read())
            {
                comboBox5.Items.Add(myReader[0]);
            }
            comboBox5.SelectedIndex = 0;
            Connection.Close();
        }
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
            SqlCommand comm1 = new SqlCommand("SELECT cena_ustanovki FROM Dopolnitelno", Connection);
            SqlDataReader myReader = comm1.ExecuteReader(CommandBehavior.CloseConnection);
            while (myReader.Read())
            {
                textBox6.Text = Convert.ToString(myReader[0]);
            }
            Connection.Close();
            try
            {
                Connection.Open();
            }
            catch (SqlException)
            {
                MessageBox.Show("Проверьте, достаточно ли места на диске, достаточно ли прав у учетной записи для операций с БД (См. справку), файлы MDF и LDF не должны быть помечены \"Только для чтения\". \n\nВозможно стоит попробовать отключить БД и запустить программу еще раз.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            SqlCommand comm2 = new SqlCommand("SELECT cena_nalichniki FROM Dopolnitelno", Connection);
            SqlDataReader myReader2 = comm2.ExecuteReader(CommandBehavior.CloseConnection);
            while (myReader2.Read())
            {
                textBox1.Text = Convert.ToString(myReader2[0]);
            }
            Connection.Close();
            try
            {
                Connection.Open();
            }
            catch (SqlException)
            {
                MessageBox.Show("Проверьте, достаточно ли места на диске, достаточно ли прав у учетной записи для операций с БД (См. справку), файлы MDF и LDF не должны быть помечены \"Только для чтения\". \n\nВозможно стоит попробовать отключить БД и запустить программу еще раз.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            SqlCommand comm3 = new SqlCommand("SELECT cena_zamki FROM Dopolnitelno", Connection);
            SqlDataReader myReader3 = comm3.ExecuteReader(CommandBehavior.CloseConnection);
            while (myReader3.Read())
            {
                textBox3.Text = Convert.ToString(myReader3[0]);
            }
            Connection.Close();
            try
            {
                Connection.Open();
            }
            catch (SqlException)
            {
                MessageBox.Show("Проверьте, достаточно ли места на диске, достаточно ли прав у учетной записи для операций с БД (См. справку), файлы MDF и LDF не должны быть помечены \"Только для чтения\". \n\nВозможно стоит попробовать отключить БД и запустить программу еще раз.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            SqlCommand comm4 = new SqlCommand("SELECT cena_ruchka FROM Dopolnitelno", Connection);
            SqlDataReader myReader4 = comm4.ExecuteReader(CommandBehavior.CloseConnection);
            while (myReader4.Read())
            {
                textBox4.Text = Convert.ToString(myReader4[0]);
            }
            Connection.Close();
            try
            {
                Connection.Open();
            }
            catch (SqlException)
            {
                MessageBox.Show("Проверьте, достаточно ли места на диске, достаточно ли прав у учетной записи для операций с БД (См. справку), файлы MDF и LDF не должны быть помечены \"Только для чтения\". \n\nВозможно стоит попробовать отключить БД и запустить программу еще раз.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            SqlCommand comm5 = new SqlCommand("SELECT cena_petli FROM Dopolnitelno", Connection);
            SqlDataReader myReader5 = comm5.ExecuteReader(CommandBehavior.CloseConnection);
            while (myReader5.Read())
            {
                textBox5.Text = Convert.ToString(myReader5[0]);
            }
            Connection.Close();
        }
        private void Calc_Load(object sender, EventArgs e)
        {
            Selpr();
            SelCena();
            ShowData();
        }
        private void button5_Click(object sender, EventArgs e)
        {
            double sum;
            sum = 0;
            textBox2.Clear();
            comboBox5.SelectedIndex = comboBox1.SelectedIndex;
            sum = Convert.ToDouble(comboBox5.Text) * Convert.ToDouble(comboBox2.Text) * Convert.ToDouble(comboBox3.Text) * Convert.ToDouble(numericUpDown1.Text) + 
                Convert.ToDouble(textBox6.Text) * Convert.ToInt32(checkBox1.Checked) + 
                Convert.ToDouble(textBox1.Text) * Convert.ToInt32(checkBox2.Checked) + 
                Convert.ToDouble(textBox3.Text) * Convert.ToInt32(checkBox3.Checked) + 
                Convert.ToDouble(textBox4.Text) * Convert.ToInt32(checkBox4.Checked) + 
                Convert.ToDouble(textBox5.Text) * Convert.ToInt32(checkBox5.Checked);
            textBox2.Text = Convert.ToString(sum);
        }
        private void SomethingChanged(object sender, EventArgs e)
        {
            CheckIsEverythingCorrect();
        }
        private void CheckIsEverythingCorrect()
        {
            if (numericUpDown1.Value > 0 && IsStringCorrectWordsOnly(comboBox1.Text) && IsStringCorrectPositiveDigitsOnly(comboBox2.Text) && IsStringCorrectPositiveDigitsOnly(comboBox3.Text))
            {
                button5.Visible = true;
            }
            else
            {
                button5.Visible = false;
            }
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
                        if (number >= 10)
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

        private void Calc_HelpRequested(object sender, HelpEventArgs hlpevent)
        {
            System.Diagnostics.Process.Start(Environment.CurrentDirectory + "\\Resources\\HelpFile.chm");
        }
    }

}
