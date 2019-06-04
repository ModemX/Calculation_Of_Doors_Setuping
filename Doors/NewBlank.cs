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
                $"1. Согласно настоящему договору Исполнитель обязуется по заданию Заказчика оказать следующие услуги: ")
                .FontSize(12);

            #region СписокТребований
            List СписокТребований = document.AddList();
            document.AddListItem(СписокТребований, " выполнить изготовление двери(-ей)", 1, ListItemType.Numbered);

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

            document.InsertList(СписокТребований);
            #endregion

            document.InsertParagraph("В свою очередь Заказчик обязуется оплатить эти услуги.\n");

            #region Калькулирование Стоимости

            int[] mass = new int[5];
            List <double> ЦенаПрофиля = new List<double>();

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
                ЦенаПрофиля.Add(Math.Round(Convert.ToDouble(myReader6[0]),4));
            }
            Connection.Close();

            double sum = ЦенаПрофиля[id_Профиля] * Высота * Ширина * Количество * + mass[0] * Convert.ToInt32(ustanovka) +
                mass[1] * Convert.ToInt32(nalichniki) +
                mass[2] * Convert.ToInt32(zamok) +
                mass[3] * Convert.ToInt32(ruchka) +
                mass[4] * Convert.ToInt32(petli);

            #endregion

            document.InsertParagraph($"2. Стоимость оказываемых услуг составляет: {sum} бел.руб.").Alignment=Alignment.both;

            document.InsertParagraph($"3. Услуги оплачиваются в течении четырнадцати календарных дней с момента заключения договора.");

            document.InsertParagraph($"4. В случае невозможности исполнения, возникшей по вине Заказчика, услуги подлежат оплате в полном объеме.");

            document.InsertParagraph($"5. В случае, когда невозможность исполнения возникла по обстоятельствам, за которые ни одна из сторон не отвечает, Заказчик возмещает Исполнителю фактически понесенные им расходы.");

            document.InsertParagraph($"6. Заказчик вправе отказаться от исполнения настоящего договора при условии оплаты Исполнителю фактически понесенных им расходов.");

            document.InsertParagraph($"7. Исполнитель вправе отказаться от исполнения настоящего договора при условии полного возмещения Заказчику убытков.");

            document.InsertParagraph($"8. К настоящему договору применяются общие положения о подряде (статьи 656 - 682 ГК РБ ) и положения о бытовом подряде (статьи 683 - 695 ГК РБ), если это не противоречит статьям 733 - 737 ГК, регулирующим вопросы возмездного оказания услуг.");

            document.InsertParagraph($"9. Срок действия настоящего договора:");

            #region Подсчет и увеличение срока заказа
            DateTime IncreasedDateOfOrdering = DateTime.Parse($"{ ДатаЗаказа.Day}.{ ДатаЗаказа.Month}.{ ДатаЗаказа.Year}");
            IncreasedDateOfOrdering.AddDays(14);
            #endregion

            document.InsertParagraph($"Начало: {DataReturner(ДатаЗаказа.Day, ДатаЗаказа.Month, ДатаЗаказа.Year)}\n" +
                $"Конец: {DataReturner(IncreasedDateOfOrdering.Day, IncreasedDateOfOrdering.Month, IncreasedDateOfOrdering.Year)}");

            document.InsertParagraph("11. Договор составлен в 2-х экземплярах, по одному для каждой из сторон.\n" +
                "12.Адреса и банковские реквизиты сторон:\n\n" +
                "Заказчик:_____________________________________________________________________________________\n\n" +
                "Исполнитель: \"ООО Наши Двери\", ул. Неизвестная 16, г.Минск. *Какие-то банковские реквизиты здесь*.\n\n\n");

            document.InsertParagraph($"Подпись Заказчика: ____________________\n");
            document.InsertParagraph($"Подпись Исполнителя: ____________________");

            document.Save();
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
    }
}