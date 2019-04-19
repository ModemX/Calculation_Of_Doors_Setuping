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
    public partial class Uchet_zakaz : Form
    {
        public Uchet_zakaz()
        {
            InitializeComponent();
        }
        public string ConnectionString = @"Data Source=.\SQLEXPRESS;AttachDbFilename=" + Application.StartupPath + @"\DOORS.mdf;Integrated Security=True;Connect Timeout=30";
        public void ShowData()
        {


            SqlConnection Connection = new SqlConnection(ConnectionString);

            Connection.Open();

            SqlCommand comm = new SqlCommand("SELECT count(id_zakaz) FROM View1", Connection);
            Int32 kol = (Int32)comm.ExecuteScalar();
            label5.Text = "всего записей - " + Convert.ToString(kol);
            if (kol > 0)
                dataGridView1.RowCount = kol + 1;
            else dataGridView1.RowCount = 1;

            SqlCommand comm1 = new SqlCommand("SELECT * FROM View1", Connection);
            SqlDataReader myReader = comm1.ExecuteReader(CommandBehavior.CloseConnection);

            int i = 0;
            while (myReader.Read())
            {
                for (int j = 0; j < 12; j++)
                {

                    dataGridView1[j, i].Value = Convert.ToString(myReader[j]);

                }
                i++;

            }


            Connection.Close();
        }        
        private void Uchet_zakaz_Load(object sender, EventArgs e)
        {
            ShowData();
        }
    }
}
