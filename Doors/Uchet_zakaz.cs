using Microsoft.Office.Interop.Excel;
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
using Excel = Microsoft.Office.Interop.Excel;

namespace Doors
{
    public partial class Uchet_zakaz : Form
    {
        public Uchet_zakaz()
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
                System.Windows.Forms.Application.Exit();
            }
            SqlCommand comm = new SqlCommand("SELECT count(id_zakaz) FROM View1", Connection);
            int kol = (int)comm.ExecuteScalar();
            label5.Text = "всего записей - " + Convert.ToString(kol);
            if (kol > 0)
            {
                dataGridView1.RowCount = kol + 1;
            }
            else
            {
                dataGridView1.RowCount = 1;
            }
            SqlCommand comm1 = new SqlCommand("SELECT * FROM View1", Connection);
            SqlDataReader myReader = comm1.ExecuteReader(CommandBehavior.CloseConnection);

            int i = 0;
            while (myReader.Read())
            {
                for (int j = 0; j < 11; j++)
                {
                    dataGridView1[j, i].Value = Convert.ToString(myReader[j + 1]);
                    if ((dataGridView1[j, i].Value as string).Contains(" 12:00:00 AM"))
                    {
                        dataGridView1[j, i].Value = (dataGridView1[j, i].Value as string).Replace(" 12:00:00 AM", "");
                    }
                }
                i++;
            }
            Connection.Close();
        }
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
                System.Windows.Forms.Application.Exit();
            }
            SqlCommand comm1 = new SqlCommand("SELECT profil FROM Profili", Connection);
            SqlDataReader myReader = comm1.ExecuteReader(CommandBehavior.CloseConnection);
            comboBox3.Items.Clear();
            while (myReader.Read())
            {
                comboBox3.Items.Add(myReader[0]);
            }
            comboBox3.SelectedIndex = 0;
            Connection.Close();
        }
        private void Uchet_zakaz_Load(object sender, EventArgs e)
        {
            ShowData();
            Selpr();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                dataGridView1.Rows[i].Selected = false;
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                    {
                        if (dataGridView1.Rows[i].Cells[1].Value.ToString() == comboBox3.Text)
                        {
                            dataGridView1.Rows[i].Selected = true;
                            break;
                        }
                    }
                }
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {

                if (Convert.ToDateTime(dataGridView1.Rows[i].Cells[0].Value + " 12:00:00 AM") < dateTimePicker1.Value || Convert.ToDateTime(dataGridView1.Rows[i].Cells[0].Value + " 12:00:00 AM") > dateTimePicker2.Value)
                {
                    dataGridView1.Rows[i].Visible = false;
                }
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                dataGridView1.Rows[i].Visible = true;
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Hide();
            Add_zakaz f2 = new Add_zakaz();
            f2.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SqlConnection Connection = new SqlConnection(ConnectionString);
            int id_zakaz = dataGridView1.CurrentCell.RowIndex + 1;
            string TextCommand = "Delete from Zakazy where id_zakaz=" + id_zakaz.ToString();
            SqlCommand Command = new SqlCommand(TextCommand, Connection);
            try
            {
                Connection.Open();
            }
            catch (SqlException)
            {
                MessageBox.Show("Проверьте, достаточно ли места на диске, достаточно ли прав у учетной записи для операций с БД (См. справку), файлы MDF и LDF не должны быть помечены \"Только для чтения\". \n\nВозможно стоит попробовать отключить БД и запустить программу еще раз.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Windows.Forms.Application.Exit();
            }
            Command.ExecuteNonQuery();
            Connection.Close();
            MessageBox.Show("данные удалены");
            ShowData();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Excel.Application exApp = new Excel.Application
            {
                Visible = true
            };
            exApp.Workbooks.Add();
            Worksheet workSheet = (Worksheet)exApp.ActiveSheet;
            workSheet.Cells[1, 1] = "Дата";
            workSheet.Cells[1, 2] = "Профиль";
            workSheet.Cells[1, 3] = "Высота (см)";
            workSheet.Cells[1, 4] = "Ширина (см)";
            workSheet.Cells[1, 5] = "Количество";
            workSheet.Cells[1, 6] = "Установка";
            workSheet.Cells[1, 7] = "Наличники";
            workSheet.Cells[1, 8] = "Замок";
            workSheet.Cells[1, 9] = "Ручка";
            workSheet.Cells[1, 10] = "Петли";
            workSheet.Cells[1, 11] = "Стоимость";
            int rowExcel = 2;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Visible == true)
                {
                    workSheet.Cells[rowExcel, "A"] = dataGridView1.Rows[i].Cells[0].Value;
                    workSheet.Cells[rowExcel, "B"] = dataGridView1.Rows[i].Cells[1].Value;
                    workSheet.Cells[rowExcel, "C"] = dataGridView1.Rows[i].Cells[2].Value;
                    workSheet.Cells[rowExcel, "D"] = dataGridView1.Rows[i].Cells[3].Value;
                    workSheet.Cells[rowExcel, "E"] = dataGridView1.Rows[i].Cells[4].Value;
                    workSheet.Cells[rowExcel, "F"] = dataGridView1.Rows[i].Cells[5].Value;
                    workSheet.Cells[rowExcel, "G"] = dataGridView1.Rows[i].Cells[6].Value;
                    workSheet.Cells[rowExcel, "H"] = dataGridView1.Rows[i].Cells[7].Value;
                    workSheet.Cells[rowExcel, "I"] = dataGridView1.Rows[i].Cells[8].Value;
                    workSheet.Cells[rowExcel, "J"] = dataGridView1.Rows[i].Cells[9].Value;
                    workSheet.Cells[rowExcel, "K"] = dataGridView1.Rows[i].Cells[10].Value;
                    ++rowExcel;
                }
            }
            workSheet.Columns.EntireColumn.AutoFit();
            string pathToXmlFile;
            pathToXmlFile = Environment.CurrentDirectory + "\\" + "Заказы.xlsx";
            workSheet.SaveAs(pathToXmlFile);
            exApp.Quit();
        }

        private void Uchet_zakaz_HelpRequested(object sender, HelpEventArgs hlpevent)
        {
            System.Diagnostics.Process.Start(Environment.CurrentDirectory + "\\Resources\\HelpFile.chm");
        }

        private void ShowStatistics_Click(object sender, EventArgs e)
        {
            Statistics statisticsWindow = new Statistics();
            statisticsWindow.ShowDialog();
        }
    }
}
