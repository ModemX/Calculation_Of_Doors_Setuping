using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Xceed.Words.NET;
using Font = Xceed.Words.NET.Font;

namespace Doors
{
    public class NewBlank
    {
        public NewBlank(int id_Заказа, DateTime ДатаЗаказа, int id_Профиля, int Высота, int Ширина, int Количество, bool ustanovka, bool nalichniki, bool zamok, bool ruchka, bool petli, string Заказчик_ФИО, int Заказчик_ТелефонныйПрефикс, int Заказчик_Телефон)
        {
            string ConnectionString = @"Data Source=.\SQLEXPRESS;AttachDbFilename=" + System.Windows.Forms.Application.StartupPath + @"\Resources\doors.mdf;Integrated Security=True;Connect Timeout=30";

            Directory.CreateDirectory(Environment.CurrentDirectory + @"\Оформленные заказы");
            string path = Environment.CurrentDirectory + @"\Оформленные заказы\Заказ №" + (100000 + id_Заказа).ToString() + ".docx";

            DocX document = DocX.Create(path);

            document.MarginTop = 30;
            document.MarginLeft = 30;
            document.MarginRight = 30;
            document.MarginBottom = 30;

            document.InsertParagraph($"{DataReturner(ДатаЗаказа.Day, ДатаЗаказа.Month, ДатаЗаказа.Year)}\n").
                     Font("Times New Roman").
                     FontSize(11).
                     Alignment = Alignment.right;

            document.InsertParagraph("Бланк-заказ на установку двери\n\n").
                     Font("Times New Roman").
                     FontSize(28).
                     Bold().
                     Alignment = Alignment.center;

            document.InsertParagraph
                ("      \"ООО Наши Двери\" именуемое в дальнейшем Исполнитель, в лице менеджера по продажам " +
                $"\"Ирина М.\" и директора компании Никита К. с одной стороны, и {Заказчик_ФИО} " +
                $"именуемое в дальнейшем Заказчик с другой стороны, заключили настоящий договор " +
                $"о нижеследующем." +
                $"\n\n" +
                $"1. Согласно настоящему договору Исполнитель обязуется по заданию Заказчика оказать следующие услуги: ").
                Font("Times New Roman").
                FontSize(12);

            #region СписокТребований
            List СписокТребований = document.AddList();
            document.AddListItem(СписокТребований, $" выполнить изготовление двери(-ей), в количестве: {Количество} экземпляр(-ов)", 1, ListItemType.Numbered);

            if (ustanovka)
            {
                document.AddListItem(СписокТребований, " выполнить установку двери(-ей)", 1, ListItemType.Numbered);
            }

            if (nalichniki)
            {
                document.AddListItem(СписокТребований, " выполнить установку наличников", 1, ListItemType.Numbered);
            }

            if (zamok)
            {
                document.AddListItem(СписокТребований, " выполнить установку замка(-ов) в дверь(-и)", 1, ListItemType.Numbered);
            }

            if (ruchka)
            {
                document.AddListItem(СписокТребований, " выполнить установку ручки (ручек) в дверь(-и)", 1, ListItemType.Numbered);
            }

            if (petli)
            {
                document.AddListItem(СписокТребований, " выполнить установку петель", 1, ListItemType.Numbered);
            }

            document.InsertList(СписокТребований, new Font("Times New Roman"), 12);
            #endregion

            document.InsertParagraph("В свою очередь Заказчик обязуется оплатить эти услуги.\n")
                .Font("Times New Roman")
                .FontSize(12);
            #region Калькулирование Стоимости

            int[] mass = new int[5];
            List<double> ЦенаПрофиля = new List<double>();

            #region Запросы к БД
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
                mass[0] = Convert.ToInt32(myReader[0]);
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
                mass[1] = Convert.ToInt32(myReader2[0]);
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
                mass[2] = Convert.ToInt32(myReader3[0]);
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
                mass[3] = Convert.ToInt32(myReader4[0]);
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
                mass[4] = Convert.ToInt32(myReader5[0]);
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
            SqlCommand comm6 = new SqlCommand("SELECT cena FROM Profili", Connection);
            SqlDataReader myReader6 = comm6.ExecuteReader(CommandBehavior.CloseConnection);
            while (myReader6.Read())
            {
                ЦенаПрофиля.Add(Math.Round(Convert.ToDouble(myReader6[0]), 4));
            }
            Connection.Close();

            #endregion

            double sum = ЦенаПрофиля[id_Профиля] * Высота * Ширина * Количество +
                mass[0] * Convert.ToInt32(ustanovka) +
                mass[1] * Convert.ToInt32(nalichniki) +
                mass[2] * Convert.ToInt32(zamok) +
                mass[3] * Convert.ToInt32(ruchka) +
                mass[4] * Convert.ToInt32(petli);

            #endregion

            document.InsertParagraph($"2. Стоимость оказываемых услуг составляет: {sum} бел.руб.")
            .Font("Times New Roman")
            .FontSize(12)
            .Alignment = Alignment.both;

            document.InsertParagraph($"3. Услуги оплачиваются в течении четырнадцати календарных дней с момента заключения договора.")
            .Font("Times New Roman")
            .FontSize(12)
            .Alignment = Alignment.both;

            document.InsertParagraph($"4. В случае невозможности исполнения, возникшей по вине Заказчика, услуги подлежат оплате в полном объеме.")
            .Font("Times New Roman")
            .FontSize(12)
            .Alignment = Alignment.both;

            document.InsertParagraph($"5. В случае, когда невозможность исполнения возникла по обстоятельствам, за которые ни одна из сторон не отвечает, Заказчик возмещает Исполнителю фактически понесенные им расходы.")
            .Font("Times New Roman")
            .FontSize(12)
            .Alignment = Alignment.both;

            document.InsertParagraph($"6. Заказчик вправе отказаться от исполнения настоящего договора при условии оплаты Исполнителю фактически понесенных им расходов.")
            .Font("Times New Roman")
            .FontSize(12)
            .Alignment = Alignment.both;

            document.InsertParagraph($"7. Исполнитель вправе отказаться от исполнения настоящего договора при условии полного возмещения Заказчику убытков.")
            .Font("Times New Roman")
            .FontSize(12)
            .Alignment = Alignment.both;

            document.InsertParagraph($"8. К настоящему договору применяются общие положения о подряде (статьи 656 - 682 ГК РБ ) и положения о бытовом подряде (статьи 683 - 695 ГК РБ), если это не противоречит статьям 733 - 737 ГК, регулирующим вопросы возмездного оказания услуг.")
            .Font("Times New Roman")
            .FontSize(12)
            .Alignment = Alignment.both;

            document.InsertParagraph($"9. Срок действия настоящего договора:")
            .Font("Times New Roman")
            .FontSize(12)
            .Alignment = Alignment.both;

            #region Подсчет и увеличение срока заказа
            DateTime IncreasedDateOfOrdering = DateTime.Parse($"{ ДатаЗаказа.Day}.{ ДатаЗаказа.Month}.{ ДатаЗаказа.Year}");
            IncreasedDateOfOrdering = IncreasedDateOfOrdering.AddDays(14);
            #endregion

            document.InsertParagraph($"Начало: {DataReturner(ДатаЗаказа.Day, ДатаЗаказа.Month, ДатаЗаказа.Year)}\n" +
                $"Конец: {DataReturner(IncreasedDateOfOrdering.Day, IncreasedDateOfOrdering.Month, IncreasedDateOfOrdering.Year)}")
            .Font("Times New Roman")
            .FontSize(12);

            document.InsertParagraph
                ("11. Договор составлен в 2-х экземплярах, по одному для каждой из сторон.")
            .Font("Times New Roman")
            .FontSize(12)
            .Alignment = Alignment.both;

            document.InsertParagraph
                ("12.Адреса и банковские реквизиты сторон:\n\n" +
                $"Заказчик:_________________________________________________________________________________________________________________________________________________________________________. Контактный телефон: +375({Заказчик_ТелефонныйПрефикс}){Заказчик_Телефон}\n\n" +
                "Исполнитель: \"ООО Наши Двери\", ул. Неизвестная 16, г.Минск. \n*Какие-то банковские реквизиты здесь*.\n\n\n")
            .Font("Times New Roman")
            .FontSize(12);

            document.InsertParagraph($"Подпись Заказчика: ____________________\n")
            .Font("Times New Roman")
            .FontSize(12);

            document.InsertParagraph($"Подпись Исполнителя: ____________________")
            .Font("Times New Roman")
            .FontSize(12);

            document.Save();
        }
        public NewBlank
            (
            System.Windows.Forms.DataVisualization.Charting.Chart chartСпросЗаКвартал,
            System.Windows.Forms.DataVisualization.Charting.Chart chartСпросЗаМесяц
            )
        {
            string ConnectionString = @"Data Source=.\SQLEXPRESS;AttachDbFilename=" + Application.StartupPath + @"\Resources\doors.mdf;Integrated Security=True;Connect Timeout=30";

            Directory.CreateDirectory(Environment.CurrentDirectory + $@"\Отчеты по продажам\{DateTime.Now.Year}\{НазваниеМесяца(DateTime.Now.Month, true)}\");
            string path = Environment.CurrentDirectory + $@"\Отчеты по продажам\{DateTime.Now.Year}\{НазваниеМесяца(DateTime.Now.Month, true)}\Отчет по продажам и предыдущий квартал ({НазваниеМесяца(DateTime.Now.Month, true)} {DateTime.Now.Year}).docx";
            SqlConnection Connection = new SqlConnection(ConnectionString);

            DocX document = DocX.Create(path);

            document.MarginTop = 30;
            document.MarginLeft = 30;
            document.MarginRight = 30;
            document.MarginBottom = 30;

            document.InsertParagraph($"Дата составления: {DataReturner(DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year)}\n").
                Font("Times New Roman").
                FontSize(11).
                Alignment = Alignment.right;

            document.InsertParagraph("Отчет о продажах за текущий месяц и предыдущий квартал\n\n").
                 Font("Times New Roman").
                 FontSize(28).
                 Bold().
                 Alignment = Alignment.center;

            document.InsertParagraph("На приведенном ниже графике отображен график спроса дверей за месяц:\n")
                .Font("Times New Roman")
                .FontSize(12)
                .Alignment = Alignment.left;

            #region Составление и заполнение графиков продаж за месяц

            PieChart WordChart_СпросЗаМесяц = new PieChart();
            WordChart_СпросЗаМесяц.AddLegend(ChartLegendPosition.Right, false);

            try
            {
                Connection.Open();
            }
            catch (SqlException)
            {
                MessageBox.Show("Проверьте, достаточно ли места на диске, достаточно ли прав у учетной записи для операций с БД (См. справку), файлы MDF и LDF не должны быть помечены \"Только для чтения\". \n\nВозможно стоит попробовать отключить БД и запустить программу еще раз.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            SqlCommand comm = new SqlCommand("SELECT count(id_profil) FROM Profili", Connection);
            object Результат = comm.ExecuteScalar();

            if (Результат != null)
            {
                int Количество_профилей = (int)Результат;

                List<string> Name = new List<string>();
                List<int> Value = new List<int>();
                Series series = new Series("Спрос товара за месяц, группировка по профилю");

                for (int i = 0; i < Количество_профилей; i++)
                {
                    comm = new SqlCommand($"SELECT profil FROM Profili where id_profil={i + 1}", Connection);
                    string НазваниеПрофиля = (string)comm.ExecuteScalar();
                    comm = new SqlCommand($"select kolvo from Zakazy where id_profil = {i + 1} and data between '{DateTime.Now.Year}/{DateTime.Now.Month}/01' and '{DateTime.Now.Year}/{DateTime.Now.Month}/{DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month)}'", Connection);
                    Результат = comm.ExecuteScalar();
                    if (Результат != null)
                    {
                        Name.Add($"{НазваниеПрофиля} ({(int)Результат})");
                        Value.Add((int)Результат);
                    }
                }
                series.Bind(Name, Value);
                WordChart_СпросЗаМесяц.AddSeries(series);
            }
            Connection.Close();
            document.InsertChart(WordChart_СпросЗаМесяц);
            #endregion

            document.InsertParagraph("\n\nА также диаграмма продаж за предыдущий квартал. Она предоставлена ниже:\n")
                .Font("Times New Roman")
                .FontSize(12)
                .Alignment = Alignment.left;

            #region Составление и заполнение графиков продаж за предыдущий квартал

            BarChart WordChart_СпросЗаКвартал = new BarChart();
            WordChart_СпросЗаКвартал.AddLegend(ChartLegendPosition.Right, false);
            WordChart_СпросЗаКвартал.BarDirection = BarDirection.Column;
            WordChart_СпросЗаКвартал.BarGrouping = BarGrouping.Clustered;
            WordChart_СпросЗаКвартал.GapWidth = 200;

            try
            {
                Connection.Open();
            }
            catch (SqlException)
            {
                MessageBox.Show("Проверьте, достаточно ли места на диске, достаточно ли прав у учетной записи для операций с БД (См. справку), файлы MDF и LDF не должны быть помечены \"Только для чтения\". \n\nВозможно стоит попробовать отключить БД и запустить программу еще раз.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            comm = new SqlCommand("SELECT count(id_profil) FROM Profili", Connection);
            Результат = comm.ExecuteScalar();

            if (Результат != null)
            {
                List<string> Name = new List<string>();
                List<int> Value = new List<int>();

                switch (Statistics.КакойКварталОтобразить())
                {
                    case 1:
                        {
                            comm = new SqlCommand($"SELECT sum(kolvo) FROM Zakazy where data between " +
                                $"'{DateTime.Now.Year - 1}/10/01' and " +
                                $"'{DateTime.Now.Year - 1}/10/{DateTime.DaysInMonth(DateTime.Now.Year - 1, 10)}'", Connection);
                            Результат = comm.ExecuteScalar();
                            if (Результат != null)
                            {
                                Name.Add($"Октябрь ({DateTime.Now.Year - 1}) ({(int)Результат})");
                                Value.Add((int)Результат);
                            }
                            else
                            {
                                Name.Add($"Октябрь ({DateTime.Now.Year - 1}) (0)");
                                Value.Add(0);
                            }


                            comm = new SqlCommand($"SELECT sum(kolvo) FROM Zakazy where data between " +
                                $"'{DateTime.Now.Year - 1}/11/01' and " +
                                $"'{DateTime.Now.Year - 1}/11/{DateTime.DaysInMonth(DateTime.Now.Year - 1, 11)}'", Connection);
                            Результат = comm.ExecuteScalar();
                            if (Результат != null)
                            {
                                Name.Add($"Ноябрь ({DateTime.Now.Year - 1}) ({(int)Результат})");
                                Value.Add((int)Результат);
                            }
                            else
                            {
                                Name.Add($"Ноябрь ({DateTime.Now.Year - 1}) (0)");
                                Value.Add(0);
                            }

                            comm = new SqlCommand($"SELECT sum(kolvo) FROM Zakazy where data between " +
                                $"'{DateTime.Now.Year - 1}/12/01' and " +
                                $"'{DateTime.Now.Year - 1}/12/{DateTime.DaysInMonth(DateTime.Now.Year - 1, 12)}'", Connection);
                            Результат = comm.ExecuteScalar();
                            if (Результат != null)
                            {
                                Name.Add($"Декабрь ({DateTime.Now.Year - 1}) ({(int)Результат})");
                                Value.Add((int)Результат);
                            }
                            else
                            {
                                Name.Add($"Декабрь ({DateTime.Now.Year - 1}) (0)");
                                Value.Add(0);
                            }
                        }
                        break;
                    case 2:
                        {
                            comm = new SqlCommand($"SELECT sum(kolvo) FROM Zakazy where data between " +
                                $"'{DateTime.Now.Year}/01/01' and " +
                                $"'{DateTime.Now.Year}/01/{DateTime.DaysInMonth(DateTime.Now.Year, 01)}'", Connection);
                            Результат = comm.ExecuteScalar();
                            if (Результат != null)
                            {
                                Name.Add($"Январь ({(int)Результат})");
                                Value.Add((int)Результат);
                            }
                            else
                            {
                                Name.Add($"Январь (0)");
                                Value.Add(0);
                            }
                            Name.Add($"Январь");
                            Value.Add((int)comm.ExecuteScalar());

                            comm = new SqlCommand($"SELECT sum(kolvo) FROM Zakazy where data between " +
                                $"'{DateTime.Now.Year}/02/01' and " +
                                $"'{DateTime.Now.Year}/02/{DateTime.DaysInMonth(DateTime.Now.Year, 02)}'", Connection);
                            Результат = comm.ExecuteScalar();
                            if (Результат != null)
                            {
                                Name.Add($"Февраль ({(int)Результат})");
                                Value.Add((int)Результат);
                            }
                            else
                            {
                                Name.Add($"Февраль (0)");
                                Value.Add(0);
                            }

                            comm = new SqlCommand($"SELECT sum(kolvo) FROM Zakazy where data between " +
                                $"'{DateTime.Now.Year}/03/01' and " +
                                $"'{DateTime.Now.Year}/03/{DateTime.DaysInMonth(DateTime.Now.Year, 03)}'", Connection);
                            Результат = comm.ExecuteScalar();
                            if (Результат != null)
                            {
                                Name.Add($"Март ({(int)Результат})");
                                Value.Add((int)Результат);
                            }
                            else
                            {
                                Name.Add($"Март (0)");
                                Value.Add(0);
                            }
                        }
                        break;
                    case 3:
                        {
                            comm = new SqlCommand($"SELECT sum(kolvo) FROM Zakazy where data between " +
                                $"'{DateTime.Now.Year}/04/01' and " +
                                $"'{DateTime.Now.Year}/04/{DateTime.DaysInMonth(DateTime.Now.Year, 04)}'", Connection);
                            Результат = comm.ExecuteScalar();
                            if (Результат != null)
                            {
                                Name.Add($"Апрель ({(int)Результат})");
                                Value.Add((int)Результат);
                            }
                            else
                            {
                                Name.Add($"Апрель (0)");
                                Value.Add(0);
                            }

                            comm = new SqlCommand($"SELECT sum(kolvo) FROM Zakazy where data between " +
                                $"'{DateTime.Now.Year}/05/01' and " +
                                $"'{DateTime.Now.Year}/05/{DateTime.DaysInMonth(DateTime.Now.Year, 05)}'", Connection);
                            Результат = comm.ExecuteScalar();
                            if (Результат != null)
                            {
                                Name.Add($"Май ({(int)Результат})");
                                Value.Add((int)Результат);
                            }
                            else
                            {
                                Name.Add($"Май (0)");
                                Value.Add(0);
                            }

                            comm = new SqlCommand($"SELECT sum(kolvo) FROM Zakazy where data between " +
                                $"'{DateTime.Now.Year}/06/01' and " +
                                $"'{DateTime.Now.Year}/06/{DateTime.DaysInMonth(DateTime.Now.Year, 06)}'", Connection);
                            Результат = comm.ExecuteScalar();
                            if (Результат != null)
                            {
                                Name.Add($"Июнь ({(int)Результат})");
                                Value.Add((int)Результат);
                            }
                            else
                            {
                                Name.Add($"Июнь (0)");
                                Value.Add(0);
                            }
                        }
                        break;
                    case 4:
                        {
                            comm = new SqlCommand($"SELECT sum(kolvo) FROM Zakazy where data between " +
                                $"'{DateTime.Now.Year}/07/01' and " +
                                $"'{DateTime.Now.Year}/07/{DateTime.DaysInMonth(DateTime.Now.Year, 07)}'", Connection);
                            Результат = comm.ExecuteScalar();
                            if (Результат != null)
                            {
                                Name.Add($"Июль ({(int)Результат})");
                                Value.Add((int)Результат);
                            }
                            else
                            {
                                Name.Add($"Июль (0)");
                                Value.Add(0);
                            }

                            comm = new SqlCommand($"SELECT sum(kolvo) FROM Zakazy where data between " +
                                $"'{DateTime.Now.Year}/08/01' and " +
                                $"'{DateTime.Now.Year}/08/{DateTime.DaysInMonth(DateTime.Now.Year, 08)}'", Connection);
                            Результат = comm.ExecuteScalar();
                            if (Результат != null)
                            {
                                Name.Add($"Август ({(int)Результат})");
                                Value.Add((int)Результат);
                            }
                            else
                            {
                                Name.Add($"Август (0)");
                                Value.Add(0);
                            }

                            comm = new SqlCommand($"SELECT sum(kolvo) FROM Zakazy where data between " +
                                $"'{DateTime.Now.Year}/09/01' and " +
                                $"'{DateTime.Now.Year}/09/{DateTime.DaysInMonth(DateTime.Now.Year, 09)}'", Connection);
                            Результат = comm.ExecuteScalar();
                            if (Результат != null)
                            {
                                Name.Add($"Сентябрь ({(int)Результат})");
                                Value.Add((int)Результат);
                            }
                            else
                            {
                                Name.Add($"Сентябрь (0)");
                                Value.Add(0);
                            }
                        }
                        break;
                }

                var series = new Series("Статистика проданных дверей за квартал");
                series.Bind(Name, Value);
                WordChart_СпросЗаКвартал.AddSeries(series);
            }
            Connection.Close();
            document.InsertChart(WordChart_СпросЗаКвартал);

            #endregion

            try
            {
                document.Save();
            }
            catch (IOException)
            {
                MessageBox.Show($"Ошибка при сохранении. \n\nНе удалось записать в файл. Убедитесь что требуемый файл ({path}) закрыт. \n\nИгнорируйте следующее сообщение");
            }
            MessageBox.Show($"Отчет создан. Он находится в директории: \"{path}\".");
        }
        private string DataReturner(int Day, int Month, int Year)
        {
            if (Day < 10 && Month < 10)
            {
                return $"0{Day}.0{Month}.{Year}";
            }
            else if (Day < 10 && Month >= 10)
            {
                return $"0{Day}.{Month}.{Year}";
            }
            else if (Day >= 10 && Month < 10)
            {
                return $"{Day}.0{Month}.{Year}";
            }
            else
            {
                return $"{Day}.{Month}.{Year}";
            }
        }
        private string НазваниеМесяца(int НомерМесяца, bool ВернутьСЗаглавной)
        {
            if (ВернутьСЗаглавной)
            {
                switch (НомерМесяца)
                {
                    case 1: return "Январь";
                    case 2: return "Февраль";
                    case 3: return "Март";
                    case 4: return "Апрель";
                    case 5: return "Май";
                    case 6: return "Июнь";
                    case 7: return "Июль";
                    case 8: return "Август";
                    case 9: return "Сентябрь";
                    case 10: return "Октябрь";
                    case 11: return "Ноябрь";
                    case 12: return "Декабрь";
                }
            }
            else
            {
                switch (НомерМесяца)
                {
                    case 1: return "январь";
                    case 2: return "февраль";
                    case 3: return "март";
                    case 4: return "апрель";
                    case 5: return "май";
                    case 6: return "июнь";
                    case 7: return "июль";
                    case 8: return "август";
                    case 9: return "сентябрь";
                    case 10: return "октябрь";
                    case 11: return "ноябрь";
                    case 12: return "декабрь";
                }
            }
            return "";
        }
    }
}