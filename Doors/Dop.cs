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
    public partial class Dop : Form
    {
        public Dop()
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
                MessageBox.Show("Проверьте, достаточно ли места на диске, достаточно ли прав у учетной записи для операций с БД, файлы MDF и LDF не должны быть помечены \"Только для чтения\". \n\nВозможно стоит попробовать отключить БД и запустить программу еще раз.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            SqlCommand comm1 = new SqlCommand("SELECT cena_ustanovki FROM Dopolnitelno", Connection);
            SqlDataReader myReader = comm1.ExecuteReader(CommandBehavior.CloseConnection);
            while (myReader.Read())
            {
                textBox1.Text = Convert.ToString(myReader[0]);

            }
            Connection.Close();
            try
            {
                Connection.Open();
            }
            catch (SqlException)
            {
                MessageBox.Show("Проверьте, достаточно ли места на диске, достаточно ли прав у учетной записи для операций с БД, файлы MDF и LDF не должны быть помечены \"Только для чтения\". \n\nВозможно стоит попробовать отключить БД и запустить программу еще раз.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            SqlCommand comm2 = new SqlCommand("SELECT cena_nalichniki FROM Dopolnitelno", Connection);
            SqlDataReader myReader2 = comm2.ExecuteReader(CommandBehavior.CloseConnection);
            while (myReader2.Read())
            {
                textBox2.Text = Convert.ToString(myReader2[0]);

            }
            Connection.Close();
            try
            {
                Connection.Open();
            }
            catch (SqlException)
            {
                MessageBox.Show("Проверьте, достаточно ли места на диске, достаточно ли прав у учетной записи для операций с БД, файлы MDF и LDF не должны быть помечены \"Только для чтения\". \n\nВозможно стоит попробовать отключить БД и запустить программу еще раз.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                MessageBox.Show("Проверьте, достаточно ли места на диске, достаточно ли прав у учетной записи для операций с БД, файлы MDF и LDF не должны быть помечены \"Только для чтения\". \n\nВозможно стоит попробовать отключить БД и запустить программу еще раз.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                MessageBox.Show("Проверьте, достаточно ли места на диске, достаточно ли прав у учетной записи для операций с БД, файлы MDF и LDF не должны быть помечены \"Только для чтения\". \n\nВозможно стоит попробовать отключить БД и запустить программу еще раз.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            SqlCommand comm5 = new SqlCommand("SELECT cena_petli FROM Dopolnitelno", Connection);
            SqlDataReader myReader5 = comm5.ExecuteReader(CommandBehavior.CloseConnection);
            while (myReader5.Read())
            {
                textBox5.Text = Convert.ToString(myReader5[0]);

            }
            Connection.Close();
        }
        private void Dop_Load(object sender, EventArgs e)
        {
            ShowData();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Enabled = true;
            textBox2.Enabled = true;
            textBox3.Enabled = true;
            textBox4.Enabled = true;
            textBox5.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string ust, nal, zam, ruc, pet;
            SqlConnection Connection = new SqlConnection(ConnectionString);
            try
            {
                Connection.Open();
            }
            catch (SqlException)
            {
                MessageBox.Show("Проверьте, достаточно ли места на диске, достаточно ли прав у учетной записи для операций с БД, файлы MDF и LDF не должны быть помечены \"Только для чтения\". \n\nВозможно стоит попробовать отключить БД и запустить программу еще раз.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            SqlCommand comm = new SqlCommand("SELECT max(id_dop) FROM Dopolnitelno", Connection);
            int max = (int)comm.ExecuteScalar();
            ust = textBox1.Text;
            nal = textBox2.Text;
            zam = textBox3.Text;
            ruc = textBox4.Text;
            pet = textBox5.Text;
            string sql = string.Format("UPDATE Dopolnitelno SET cena_ustanovki='{0}', cena_nalichniki='{1}', cena_zamki='{2}', cena_ruchka='{3}', cena_petli='{4}' WHERE id_dop=1", ust, nal, zam, ruc, pet);
            SqlCommand cmd = new SqlCommand(sql, Connection);
            if (cmd.ExecuteNonQuery() == 1)
            {
                MessageBox.Show("Данные успешно изменены.");
            }
            ShowData();
            Connection.Close();
            textBox1.Enabled = false;
            textBox2.Enabled = false;
            textBox3.Enabled = false;
            textBox4.Enabled = false;
            textBox5.Enabled = false;
        }
    }
}
