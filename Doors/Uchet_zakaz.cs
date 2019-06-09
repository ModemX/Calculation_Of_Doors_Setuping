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
            Работа_с_запясями.Text = "Работа с записями (Всего записей: " + Convert.ToString(kol) + ")";
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
                    if ((dataGridView1[j, i].Value as string).Contains(" 0:00:00"))
                    {
                        dataGridView1[j, i].Value = (dataGridView1[j, i].Value as string).Replace(" 0:00:00", "");
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
            for (int index = 0; index < dataGridView1.Rows.Count - 1; ++index)
            {
                if (dataGridView1.Rows[index].Cells[1].Value.ToString() != comboBox3.Text)
                    dataGridView1.Rows[index].Visible = false;
            }
            button2.Enabled = false;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                if (Convert.ToDateTime(dataGridView1.Rows[i].Cells[0].Value + " 0:00:00") < dateTimePicker1.Value || Convert.ToDateTime(dataGridView1.Rows[i].Cells[0].Value + " 0:00:00") > dateTimePicker2.Value)
                {
                    dataGridView1.Rows[i].Visible = false;
                }
            }
            button9.Enabled = false;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            ShowData();
            Selpr();
            button9.Enabled = true;
            button2.Enabled = true;
            button5.Enabled = true;
            ПоискПоКолву_От.Enabled = true;
            ПоискПоКолву_До.Enabled = true;
            ПоискПоЦене_От.Enabled = true;
            ПоискПоЦене_До.Enabled = true;
            button8.Enabled = true;
            button6.Enabled = true;
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

        private void Button5_Click(object sender, EventArgs e)
        {
            bool flag = ПоискПоКолву_От.Value <= ПоискПоКолву_До.Value;
            if (flag)
            {
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    bool flag2 = Convert.ToInt32(dataGridView1.Rows[i].Cells[4].Value) > Convert.ToInt32(ПоискПоКолву_До.Value) || Convert.ToInt32(dataGridView1.Rows[i].Cells[4].Value) < Convert.ToInt32(ПоискПоКолву_От.Value);
                    if (flag2)
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }
                }
                button5.Enabled = false;
                ПоискПоКолву_От.Enabled = false;
                ПоискПоКолву_До.Enabled = false;
            }
        }
        private void Button6_Click(object sender, EventArgs e)
        {
            SqlConnection Connection = new SqlConnection(ConnectionString);
            SqlCommand comm = new SqlCommand(string.Format("Select data, profil, vysota, shirina, kolvo, ustanovka, nalichnik, zamki, ruchki, petl, stoimost from View1 where id_zakaz = {0}", numericUpDown1.Value - 100000), Connection);
            try
            {
                Connection.Open();
            }
            catch (SqlException)
            {
                MessageBox.Show("Проверьте, достаточно ли места на диске, достаточно ли прав у учетной записи для операций с БД (См. справку), файлы MDF и LDF не должны быть помечены \"Только для чтения\". \n\nВозможно стоит попробовать отключить БД и запустить программу еще раз.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                //Application.Exit();
            }
            SqlDataReader myReader = comm.ExecuteReader(CommandBehavior.CloseConnection);
            dataGridView1.Rows.Clear();
            dataGridView1.Rows.Add();
            while (myReader.Read())
            {
                for (int i = 0; i < 11; i++)
                {
                    dataGridView1[i, 0].Value = Convert.ToString(myReader.GetValue(i));
                    bool flag = (dataGridView1[i, 0].Value as string).Contains(" 0:00:00");
                    if (flag)
                    {
                        dataGridView1[i, 0].Value = (dataGridView1[i, 0].Value as string).Replace(" 0:00:00", "");
                    }
                }
            }
            Connection.Close();
            button9.Enabled = false;
            button8.Enabled = false;
            button6.Enabled = false;
            button2.Enabled = false;
            button5.Enabled = false;
            ПоискПоКолву_От.Enabled = false;
            ПоискПоКолву_До.Enabled = false;
            ПоискПоЦене_От.Enabled = false;
            ПоискПоЦене_До.Enabled = false;
        }

        private void Button8_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                bool flag = Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value) > Convert.ToDouble(ПоискПоЦене_До.Text) || Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value) < Convert.ToDouble(ПоискПоЦене_От.Text);
                if (flag)
                {
                    dataGridView1.Rows[i].Visible = false;
                }
            }
            button8.Enabled = false;
        }

        private void ПоискПоКолву_SomethingChanged(object sender, EventArgs e)
        {
            bool flag = ПоискПоКолву_От.Value > ПоискПоКолву_До.Value;
            if (flag)
            {
                button5.Enabled = false;
            }
            else
            {
                button5.Enabled = true;
            }
        }
        private void ПоискПоЦене_SomethingChanged(object sender, EventArgs e)
        {
            string temp = "";
            double ЦенаОт = 0.0;
            double ЦенаДо = 0.0;
            bool Is_ЦенаОт_correct = false;
            for (int i = 0; i < ПоискПоЦене_От.Text.Length; i++)
            {
                char symbol = ПоискПоЦене_От.Text[i];
                bool flag = char.IsDigit(symbol) || symbol == ',' || symbol == '.';
                if (!flag)
                {
                    button8.Enabled = false;
                    Is_ЦенаОт_correct = false;
                    break;
                }
                bool flag2 = symbol == '.';
                if (flag2)
                {
                    temp += ",";
                }
                else
                {
                    temp += symbol.ToString();
                }
                bool flag3 = i == ПоискПоЦене_От.Text.Length - 1;
                if (flag3)
                {
                    bool flag4 = !double.TryParse(temp, out ЦенаОт);
                    if (flag4)
                    {
                        button8.Enabled = false;
                        Is_ЦенаОт_correct = false;
                        break;
                    }
                    Is_ЦенаОт_correct = true;
                    temp = "";
                }
            }
            bool flag5 = Is_ЦенаОт_correct;
            if (flag5)
            {
                for (int j = 0; j < ПоискПоЦене_До.Text.Length; j++)
                {
                    char symbol2 = ПоискПоЦене_До.Text[j];
                    bool flag6 = char.IsDigit(symbol2) || symbol2 == ',' || symbol2 == '.';
                    if (!flag6)
                    {
                        button8.Enabled = false;
                        break;
                    }
                    bool flag7 = symbol2 == '.';
                    if (flag7)
                    {
                        temp += ",";
                    }
                    else
                    {
                        temp += symbol2.ToString();
                    }
                    bool flag8 = j == ПоискПоЦене_До.Text.Length - 1;
                    if (flag8)
                    {
                        bool flag9 = !double.TryParse(temp, out ЦенаДо);
                        if (flag9)
                        {
                            button8.Enabled = false;
                            break;
                        }
                        bool flag10 = ЦенаОт > ЦенаДо;
                        if (flag10)
                        {
                            button8.Enabled = false;
                        }
                        else
                        {
                            button8.Enabled = true;
                        }
                    }
                }
            }

        }
    }
}
