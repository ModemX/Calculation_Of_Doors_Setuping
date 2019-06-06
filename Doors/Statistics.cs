using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Doors
{
    public partial class Statistics : Form
    {
        public Statistics()
        {
            InitializeComponent();

            Initialize_chartСпросЗаМесяц();
            Initialize_chartСпросЗаКвартал();
        }

        private void Initialize_chartСпросЗаМесяц()
        {
            string ConnectionString = @"Data Source=.\SQLEXPRESS;AttachDbFilename=" + System.Windows.Forms.Application.StartupPath + @"\Resources\doors.mdf;Integrated Security=True;Connect Timeout=30";
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
            int Количество_профилей = (int)comm.ExecuteScalar();
            if (Количество_профилей > 0)
            {
                for (int i = 0; i < Количество_профилей; i++)
                {
                    comm = new SqlCommand($"SELECT profil FROM Profili where id_profil={i + 1}", Connection);
                    string НазваниеПрофиля = (string)comm.ExecuteScalar();
                    comm = new SqlCommand($"select kolvo from Zakazy where id_profil = {i + 1} and data between '{DateTime.Now.Year}/{DateTime.Now.Month}/01' and '{DateTime.Now.Year}/{DateTime.Now.Month}/{DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month)}'", Connection);
                    object Результат = comm.ExecuteScalar();
                    int КоличествоДверейПоПрофилю = 0;
                    if (Результат != null)
                    {
                        КоличествоДверейПоПрофилю = (int)Результат;
                        chartСпросЗаМесяц.Series[0].Points.Add(КоличествоДверейПоПрофилю);
                        chartСпросЗаМесяц.Series[0].Points[i].Label = КоличествоДверейПоПрофилю.ToString();
                        chartСпросЗаМесяц.Series[0].Points[i].LegendText = НазваниеПрофиля;
                    }
                    else
                    {
                        chartСпросЗаМесяц.Series[0].Points.Add(КоличествоДверейПоПрофилю);
                        chartСпросЗаМесяц.Series[0].Points[i].Label = КоличествоДверейПоПрофилю.ToString();
                        chartСпросЗаМесяц.Series[0].Points[i].LegendText = НазваниеПрофиля;
                        chartСпросЗаМесяц.Series[0].Points[i].IsEmpty = true;
                    }
                }
            }
            Connection.Close();

            for (int i = 0; i < chartСпросЗаМесяц.Series[0].Points.Count; i++)
            {
                System.Windows.Forms.DataVisualization.Charting.DataPoint item = (System.Windows.Forms.DataVisualization.Charting.DataPoint)chartСпросЗаМесяц.Series[0].Points[i];
                if (item.IsEmpty == true)
                {
                    if (i == chartСпросЗаМесяц.Series[0].Points.Count - 1)
                    {
                        chartСпросЗаМесяц.Series[0].Points.Clear();
                        chartСпросЗаМесяц.Series[0].Points.Add(1);
                        chartСпросЗаМесяц.Series[0].Points[0].Label = "Нет данных \nдля отображения";
                        chartСпросЗаМесяц.Series[0].Points[0].IsVisibleInLegend = false;
                    }
                }
                else
                {
                    break;
                }
            }
        }

        private void Initialize_chartСпросЗаКвартал()
        {
            string ConnectionString = @"Data Source=.\SQLEXPRESS;AttachDbFilename=" + System.Windows.Forms.Application.StartupPath + @"\Resources\doors.mdf;Integrated Security=True;Connect Timeout=30";
            SqlConnection Connection = new SqlConnection(ConnectionString);
            try
            {
                Connection.Open();
            }
            catch (SqlException)
            {
                MessageBox.Show("Проверьте, достаточно ли места на диске, достаточно ли прав у учетной записи для операций с БД (См. справку), файлы MDF и LDF не должны быть помечены \"Только для чтения\". \n\nВозможно стоит попробовать отключить БД и запустить программу еще раз.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            SqlCommand comm;
            int[] КоличествоПроданныхДверейПоМесяцам = new int[3];

            switch (КакойКварталОтобразить())
            {
                case 1:
                    {
                        comm = new SqlCommand($"SELECT sum(kolvo) FROM Zakazy where data between " +
                            $"'{DateTime.Now.Year - 1}/10/01' and " +
                            $"'{DateTime.Now.Year - 1}/10/{DateTime.DaysInMonth(DateTime.Now.Year - 1, 10)}'", Connection);
                        КоличествоПроданныхДверейПоМесяцам[0] = (int)comm.ExecuteScalar();

                        comm = new SqlCommand($"SELECT sum(kolvo) FROM Zakazy where data between " +
                            $"'{DateTime.Now.Year - 1}/11/01' and " +
                            $"'{DateTime.Now.Year - 1}/11/{DateTime.DaysInMonth(DateTime.Now.Year - 1, 11)}'", Connection);
                        КоличествоПроданныхДверейПоМесяцам[1] = (int)comm.ExecuteScalar();

                        comm = new SqlCommand($"SELECT sum(kolvo) FROM Zakazy where data between " +
                            $"'{DateTime.Now.Year - 1}/12/01' and " +
                            $"'{DateTime.Now.Year - 1}/12/{DateTime.DaysInMonth(DateTime.Now.Year - 1, 12)}'", Connection);
                        КоличествоПроданныхДверейПоМесяцам[2] = (int)comm.ExecuteScalar();
                    }
                    break;
                case 2:
                    {
                        comm = new SqlCommand($"SELECT sum(kolvo) FROM Zakazy where data between " +
                            $"'{DateTime.Now.Year}/01/01' and " +
                            $"'{DateTime.Now.Year}/01/{DateTime.DaysInMonth(DateTime.Now.Year, 01)}'", Connection);
                        КоличествоПроданныхДверейПоМесяцам[0] = (int)comm.ExecuteScalar();
                        comm = new SqlCommand($"SELECT sum(kolvo) FROM Zakazy where data between " +
                            $"'{DateTime.Now.Year}/02/01' and " +
                            $"'{DateTime.Now.Year}/02/{DateTime.DaysInMonth(DateTime.Now.Year, 02)}'", Connection);
                        КоличествоПроданныхДверейПоМесяцам[1] = (int)comm.ExecuteScalar();
                        comm = new SqlCommand($"SELECT sum(kolvo) FROM Zakazy where data between " +
                            $"'{DateTime.Now.Year}/03/01' and " +
                            $"'{DateTime.Now.Year}/03/{DateTime.DaysInMonth(DateTime.Now.Year, 03)}'", Connection);
                        КоличествоПроданныхДверейПоМесяцам[2] = (int)comm.ExecuteScalar();
                    }
                    break;
                case 3:
                    {
                        comm = new SqlCommand($"SELECT sum(kolvo) FROM Zakazy where data between " +
                            $"'{DateTime.Now.Year}/04/01' and " +
                            $"'{DateTime.Now.Year}/04/{DateTime.DaysInMonth(DateTime.Now.Year, 04)}'", Connection);
                        КоличествоПроданныхДверейПоМесяцам[0] = (int)comm.ExecuteScalar();

                        comm = new SqlCommand($"SELECT sum(kolvo) FROM Zakazy where data between " +
                            $"'{DateTime.Now.Year}/05/01' and " +
                            $"'{DateTime.Now.Year}/05/{DateTime.DaysInMonth(DateTime.Now.Year, 05)}'", Connection);
                        КоличествоПроданныхДверейПоМесяцам[1] = (int)comm.ExecuteScalar();

                        comm = new SqlCommand($"SELECT sum(kolvo) FROM Zakazy where data between " +
                            $"'{DateTime.Now.Year}/06/01' and " +
                            $"'{DateTime.Now.Year}/06/{DateTime.DaysInMonth(DateTime.Now.Year, 06)}'", Connection);
                        КоличествоПроданныхДверейПоМесяцам[2] = (int)comm.ExecuteScalar();
                    }
                    break;
                case 4:
                    {
                        comm = new SqlCommand($"SELECT sum(kolvo) FROM Zakazy where data between " +
                            $"'{DateTime.Now.Year}/07/01' and " +
                            $"'{DateTime.Now.Year}/07/{DateTime.DaysInMonth(DateTime.Now.Year, 07)}'", Connection);
                        КоличествоПроданныхДверейПоМесяцам[0] = (int)comm.ExecuteScalar();

                        comm = new SqlCommand($"SELECT sum(kolvo) FROM Zakazy where data between " +
                            $"'{DateTime.Now.Year}/08/01' and " +
                            $"'{DateTime.Now.Year}/08/{DateTime.DaysInMonth(DateTime.Now.Year, 08)}'", Connection);
                        КоличествоПроданныхДверейПоМесяцам[1] = (int)comm.ExecuteScalar();

                        comm = new SqlCommand($"SELECT sum(kolvo) FROM Zakazy where data between " +
                            $"'{DateTime.Now.Year}/09/01' and " +
                            $"'{DateTime.Now.Year}/09/{DateTime.DaysInMonth(DateTime.Now.Year, 09)}'", Connection);
                        КоличествоПроданныхДверейПоМесяцам[2] = (int)comm.ExecuteScalar();
                    }
                    break;
            }

            switch (КакойКварталОтобразить())
            {
                case 1:
                    {
                        chartСпросЗаКвартал.Series[0].Points.Add(1, КоличествоПроданныхДверейПоМесяцам[0]);
                        chartСпросЗаКвартал.Series[0].Points[0].Label = КоличествоПроданныхДверейПоМесяцам[0].ToString();
                        chartСпросЗаКвартал.Series[0].Points.Add(2, КоличествоПроданныхДверейПоМесяцам[1]);
                        chartСпросЗаКвартал.Series[0].Points[1].Label = КоличествоПроданныхДверейПоМесяцам[1].ToString();
                        chartСпросЗаКвартал.Series[0].Points.Add(3, КоличествоПроданныхДверейПоМесяцам[2]);
                        chartСпросЗаКвартал.Series[0].Points[2].Label = КоличествоПроданныхДверейПоМесяцам[2].ToString();
                    }
                    break;
                case 2:
                    {
                        chartСпросЗаКвартал.Series[0].Points.Add(4, КоличествоПроданныхДверейПоМесяцам[0]);
                        chartСпросЗаКвартал.Series[0].Points[0].Label = КоличествоПроданныхДверейПоМесяцам[0].ToString();
                        chartСпросЗаКвартал.Series[0].Points.Add(5, КоличествоПроданныхДверейПоМесяцам[1]);
                        chartСпросЗаКвартал.Series[0].Points[1].Label = КоличествоПроданныхДверейПоМесяцам[1].ToString();
                        chartСпросЗаКвартал.Series[0].Points.Add(6, КоличествоПроданныхДверейПоМесяцам[2]);
                        chartСпросЗаКвартал.Series[0].Points[2].Label = КоличествоПроданныхДверейПоМесяцам[2].ToString();
                    }
                    break;
                case 3:
                    {
                        chartСпросЗаКвартал.Series[0].Points.Add(7, КоличествоПроданныхДверейПоМесяцам[0]);
                        chartСпросЗаКвартал.Series[0].Points[0].Label = КоличествоПроданныхДверейПоМесяцам[0].ToString();
                        chartСпросЗаКвартал.Series[0].Points.Add(8, КоличествоПроданныхДверейПоМесяцам[1]);
                        chartСпросЗаКвартал.Series[0].Points[1].Label = КоличествоПроданныхДверейПоМесяцам[1].ToString();
                        chartСпросЗаКвартал.Series[0].Points.Add(9, КоличествоПроданныхДверейПоМесяцам[2]);
                        chartСпросЗаКвартал.Series[0].Points[2].Label = КоличествоПроданныхДверейПоМесяцам[2].ToString();
                    }
                    break;
                case 4:
                    {
                        chartСпросЗаКвартал.Series[0].Points.Add(10, КоличествоПроданныхДверейПоМесяцам[0]);
                        chartСпросЗаКвартал.Series[0].Points[0].Label = КоличествоПроданныхДверейПоМесяцам[0].ToString();
                        chartСпросЗаКвартал.Series[0].Points.Add(11, КоличествоПроданныхДверейПоМесяцам[1]);
                        chartСпросЗаКвартал.Series[0].Points[1].Label = КоличествоПроданныхДверейПоМесяцам[1].ToString();
                        chartСпросЗаКвартал.Series[0].Points.Add(12, КоличествоПроданныхДверейПоМесяцам[2]);
                        chartСпросЗаКвартал.Series[0].Points[2].Label = КоличествоПроданныхДверейПоМесяцам[2].ToString();
                    }
                    break;
            }
        }

        private int КакойКварталОтобразить()
        {
            DateTime DateTimeNow = DateTime.Now;

            DateTime ПервыйКвартал = DateTime.Parse($"01.01.{DateTime.Now.Year}");
            DateTime ВторойКвартал = DateTime.Parse($"01.04.{DateTime.Now.Year}");
            DateTime ТретийКвартал = DateTime.Parse($"01.07.{DateTime.Now.Year}");
            DateTime ЧетвертыйКвартал = DateTime.Parse($"01.10.{DateTime.Now.Year}");

            if (DateTimeNow >= ПервыйКвартал && DateTimeNow < ВторойКвартал)
            {
                return 4;
            }
            else if (DateTimeNow >= ВторойКвартал && DateTimeNow < ТретийКвартал)
            {
                return 1;
            }
            else if (DateTimeNow >= ТретийКвартал && DateTimeNow < ЧетвертыйКвартал)
            {
                return 2;
            }
            else
            {
                return 3;
            }
        }
    }
}
