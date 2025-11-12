using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.Data.SqlClient;
using System.Net.NetworkInformation;
using System.Xml;
using System.Drawing.Printing;
using Npgsql;
using ExcelLibrary;
using ExcelLibrary.SpreadSheet;
using ExcelLibrary.CompoundDocumentFormat;
using ExcelLibrary.BinaryFileFormat;

namespace ascue_kompas
{
    public partial class Options : Form
    {
        private View options;
        
        public Options(View o)
        {
            InitializeComponent();
            options = o;
        }
        class mails
        {
            public string name_tii { get; set; }
            public string num_sch { get; set; }

        }

        public string put = "Server=192.168.150.40;Port=5432;;User Id=psql; Password=;Database=monitor_options; CommandTimeout=955555";
        public string status = "";

        List<mails> ls = new List<mails>();
        
        private void button1_Click(object sender, EventArgs e)
        {
            this.dataGridView1.Rows.Clear();
        
            //this.Text = options.comboBox1.SelectedItem.ToString();
            for (int i = 0; i < options.checkedListBox1.Items.Count; i++)
            {
                this.dataGridView1.Rows.Add(1);
                //OleDbConnection connection = new OleDbConnection(conStr);
                NpgsqlConnection connection = new NpgsqlConnection(put);
                connection.Open();


                //OleDbCommand command1 = new OleDbCommand("select count(*) from balans where name_rp = '" + form3.comboBox1.SelectedItem.ToString() + "' and name_tii = '" + form3.checkedListBox1.Items[i].ToString() + "'", connection);
                NpgsqlCommand command1 = new NpgsqlCommand("select count(*) from \"balans\" where \"name_rp\" = '" + options.comboBox1.SelectedItem.ToString() + "' and \"name_tii\" = '" + options.checkedListBox1.Items[i].ToString() + "'", connection);

                int p = Convert.ToInt32(command1.ExecuteScalar());
                if (p == 0)
                {

                    this.dataGridView1.Rows[i].Cells[0].Value = false;
                    this.dataGridView1.Rows[i].Cells[1].Value = false;
                    this.dataGridView1.Rows[i].Cells[2].Value = false;
                    this.dataGridView1.Rows[i].Cells[3].Value = false;

                }
                else
                {

                    //OleDbCommand com = new OleDbCommand("select vvod, fider, ktt from balans where name_rp = '" + form3.comboBox1.SelectedItem.ToString() + "' and name_tii = '" + form3.checkedListBox1.Items[i].ToString() + "'", connection);
                    //OleDbDataReader datar = com.ExecuteReader();
                    NpgsqlCommand com = new NpgsqlCommand("select \"vvod\", \"fider\", \"vvod_niz\", \"fider_niz\", \"ktt\" from \"balans\" where \"name_rp\" = '" + options.comboBox1.SelectedItem.ToString() + "' and \"name_tii\" = '" + options.checkedListBox1.Items[i].ToString() + "'", connection);

                    NpgsqlDataReader datar = com.ExecuteReader();

                    List<string> list = new List<string>();

                    while (datar.Read())
                    {
                        list.Add(datar.GetValue(0).ToString());
                        list.Add(datar.GetValue(1).ToString());
                        list.Add(datar.GetValue(2).ToString());
                        list.Add(datar.GetValue(3).ToString());
                        list.Add(datar.GetValue(4).ToString());

                    }
                    this.dataGridView1.Rows[i].Cells[0].Value = list[0];
                    this.dataGridView1.Rows[i].Cells[1].Value = list[1];
                    this.dataGridView1.Rows[i].Cells[2].Value = list[2];
                    this.dataGridView1.Rows[i].Cells[3].Value = list[3];
                    this.dataGridView1.Rows[i].Cells[5].Value = list[4];
                }
                connection.Close();
                this.dataGridView1.Rows[i].Cells[4].Value = options.checkedListBox1.Items[i].ToString();
                //dataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
                //dataGridView1.RowHeadersWidth = 20;
              
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            int col = e.ColumnIndex;
            int row = e.RowIndex;
            if (row > 0)
            {
                if
            (this.dataGridView1.Rows[row].Cells[0].Value != null)
                {
                    if (col == 0)
                    {

                        if (this.dataGridView1.Rows[row].Cells[0].Value.ToString().Equals("True"))
                        {
                            this.dataGridView1.Rows[row].Cells[1].Value = false;
                            this.dataGridView1.Rows[row].Cells[2].Value = false;
                            this.dataGridView1.Rows[row].Cells[3].Value = false;
                        }
                    }
                    if (col == 1)
                    {

                        if (this.dataGridView1.Rows[row].Cells[1].Value.ToString().Equals("True"))
                        {
                            this.dataGridView1.Rows[row].Cells[0].Value = false;
                            this.dataGridView1.Rows[row].Cells[2].Value = false;
                            this.dataGridView1.Rows[row].Cells[3].Value = false;
                        }
                    }
                    if (col == 2)
                    {

                        if (this.dataGridView1.Rows[row].Cells[2].Value.ToString().Equals("True"))
                        {
                            this.dataGridView1.Rows[row].Cells[0].Value = false;
                            this.dataGridView1.Rows[row].Cells[1].Value = false;
                            this.dataGridView1.Rows[row].Cells[3].Value = false;
                        }
                    }
                    if (col == 3)
                    {

                        if (this.dataGridView1.Rows[row].Cells[3].Value.ToString().Equals("True"))
                        {
                            this.dataGridView1.Rows[row].Cells[0].Value = false;
                            this.dataGridView1.Rows[row].Cells[1].Value = false;
                            this.dataGridView1.Rows[row].Cells[2].Value = false;
                        }
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DateTime data = DateTime.Now;  //Время запуска программы
            String host = System.Net.Dns.GetHostName(); //
            System.Net.IPAddress ip = System.Net.Dns.GetHostByName(host).AddressList[0];

            NpgsqlConnection connection = new NpgsqlConnection(put);
            connection.Open();
            for (int i = 0; i < this.dataGridView1.Rows.Count; i++)
            {

                if (this.dataGridView1.Rows[i].Cells[5].Value == null) { this.dataGridView1.Rows[i].Cells[5].Value = 1; }
                // OleDbCommand command1 = new OleDbCommand("select count(*) from balans where name_rp = '" + form3.comboBox1.SelectedItem.ToString() + "' and name_tii = '" + this.dataGridView1.Rows[i].Cells[2].Value.ToString() + "'", connection);
                NpgsqlCommand command1 = new NpgsqlCommand("select count(*) from \"balans\" where \"name_rp\" = '" + options.comboBox1.SelectedItem.ToString() + "' and \"name_tii\" = '" + this.dataGridView1.Rows[i].Cells[4].Value.ToString() + "'", connection);
                int p = Convert.ToInt32(command1.ExecuteScalar());
                NpgsqlCommand command6 = new NpgsqlCommand("select count(*) from \"journal\" where \"name_rp\" = '" + options.comboBox1.SelectedItem.ToString() + "' and \"name_tii\" = '" + this.dataGridView1.Rows[i].Cells[4].Value.ToString() + "'", connection);
                int p1 = Convert.ToInt32(command6.ExecuteScalar());


                if (p == 0)
                {
                    //OleDbCommand command = new OleDbCommand("insert into balans values(" + this.dataGridView1.Rows[i].Cells[0].Value.ToString() + ", " + this.dataGridView1.Rows[i].Cells[1].Value.ToString() + ", '" + form3.comboBox1.SelectedItem.ToString() + "', '" + this.dataGridView1.Rows[i].Cells[2].Value.ToString() + "', '" + this.dataGridView1.Rows[i].Cells[3].Value.ToString() + "')", connection);
                    NpgsqlCommand command = new NpgsqlCommand("insert into \"balans\" values(" + this.dataGridView1.Rows[i].Cells[0].Value.ToString() + ", " + this.dataGridView1.Rows[i].Cells[1].Value.ToString() + ", " + this.dataGridView1.Rows[i].Cells[2].Value.ToString() + ", " + this.dataGridView1.Rows[i].Cells[3].Value.ToString() + ", '" + options.comboBox1.SelectedItem.ToString() + "', '" + this.dataGridView1.Rows[i].Cells[4].Value.ToString() + "', '" + this.dataGridView1.Rows[i].Cells[5].Value.ToString() + "')", connection);

                    command.ExecuteNonQuery();

                }

                if (p1 == 0)
                {
                    NpgsqlCommand command4 = new NpgsqlCommand("insert into \"journal\" values('" + data.ToString() + "', '" + ip.ToString() + "', '" + options.comboBox1.SelectedItem.ToString() + "', '" + this.dataGridView1.Rows[i].Cells[4].Value.ToString() + "', " + this.dataGridView1.Rows[i].Cells[0].Value.ToString() + ", " + this.dataGridView1.Rows[i].Cells[1].Value.ToString() + ", '" + this.dataGridView1.Rows[i].Cells[5].Value.ToString() + "')", connection);
                    command4.ExecuteNonQuery();
                    NpgsqlCommand command7 = new NpgsqlCommand("delete from \"journal\" where \"name_tii\" like 'P%' or \"name_tii\" like 'с%' or \"name_tii\" like 'С%'", connection);
                    command7.ExecuteNonQuery();

                    //MessageBox.Show("insert into \"journal\" values('" + data.ToString() + "', '" + ip.ToString() + "', '" + form3.comboBox1.SelectedItem.ToString() + "', '" + this.dataGridView1.Rows[i].Cells[2].Value.ToString() + "', " + this.dataGridView1.Rows[i].Cells[0].Value.ToString() + ", " + this.dataGridView1.Rows[i].Cells[1].Value.ToString() + ", '" + this.dataGridView1.Rows[i].Cells[3].Value.ToString() + "')");
                }
                if (p1 != 0)
                {
                    NpgsqlCommand command4 = new NpgsqlCommand("insert into \"journal\" values('" + data.ToString() + "', '" + ip.ToString() + "', '" + options.comboBox1.SelectedItem.ToString() + "', '" + this.dataGridView1.Rows[i].Cells[4].Value.ToString() + "', " + this.dataGridView1.Rows[i].Cells[0].Value.ToString() + ", " + this.dataGridView1.Rows[i].Cells[1].Value.ToString() + ", '" + this.dataGridView1.Rows[i].Cells[5].Value.ToString() + "')", connection);
                    command4.ExecuteNonQuery();
                    NpgsqlCommand command8 = new NpgsqlCommand("delete from \"journal\" where \"name_tii\" like 'P%' or \"name_tii\" like 'с%' or \"name_tii\" like 'С%'", connection);
                    command8.ExecuteNonQuery();
                    //MessageBox.Show("insert into \"journal\" values('" + data.ToString() + "', '" + ip.ToString() + "', '" + form3.comboBox1.SelectedItem.ToString() + "', '" + this.dataGridView1.Rows[i].Cells[2].Value.ToString() + "', " + this.dataGridView1.Rows[i].Cells[0].Value.ToString() + ", " + this.dataGridView1.Rows[i].Cells[1].Value.ToString() + ", '" + this.dataGridView1.Rows[i].Cells[3].Value.ToString() + "')");
                }

                if (p == 1)
                {

                    //OleDbCommand command3 = new OleDbCommand("update balans set vvod = " + this.dataGridView1.Rows[i].Cells[0].Value.ToString() + ", fider = " + this.dataGridView1.Rows[i].Cells[1].Value.ToString() + ", ktt = " + this.dataGridView1.Rows[i].Cells[3].Value.ToString() + " where name_rp = '" + form3.comboBox1.SelectedItem.ToString() + "' and name_tii = '" + this.dataGridView1.Rows[i].Cells[2].Value.ToString() + "'", connection);
                    NpgsqlCommand command3 = new NpgsqlCommand("update \"balans\" set \"vvod\" = " + this.dataGridView1.Rows[i].Cells[0].Value.ToString() + ", \"fider\" = " + this.dataGridView1.Rows[i].Cells[1].Value.ToString() + ", \"vvod_niz\" = " + this.dataGridView1.Rows[i].Cells[2].Value.ToString() + ", \"fider_niz\" = " + this.dataGridView1.Rows[i].Cells[3].Value.ToString() + ", \"ktt\" = " + this.dataGridView1.Rows[i].Cells[5].Value.ToString() + " where \"name_rp\" = '" + options.comboBox1.SelectedItem.ToString() + "' and \"name_tii\" = '" + this.dataGridView1.Rows[i].Cells[4].Value.ToString() + "'", connection);
                    command3.ExecuteNonQuery();
                }



            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.dataGridView2.Rows.Clear();
            dataGridView3.DataSource = null;
            //this.Text = options.comboBox1.SelectedItem.ToString();
            for (int i = 0; i < options.checkedListBox1.Items.Count; i++)
            {
                this.dataGridView2.Rows.Add(1);
                //OleDbConnection connection = new OleDbConnection(conStr);
                NpgsqlConnection connection = new NpgsqlConnection(put);

                connection.Open();


                // OleDbCommand command1 = new OleDbCommand("select count(*) from number_sch where name_rp = '" + form3.comboBox1.SelectedItem.ToString() + "' and name_tii = '" + form3.checkedListBox1.Items[i].ToString() + "'", connection);
                NpgsqlCommand command1 = new NpgsqlCommand("select count(*) from \"number_sch\" where \"name_rp\" = '" + options.comboBox1.SelectedItem.ToString() + "' and \"name_tii\" = '" + options.checkedListBox1.Items[i].ToString() + "'", connection);

                int p = Convert.ToInt32(command1.ExecuteScalar());
                if (p == 0)
                {

                    // this.dataGridView2.Rows[i].Cells[0].Value = null;
                    //this.dataGridView2.Rows[i].Cells[1].Value = null;

                }
                else
                {

                    //OleDbCommand com = new OleDbCommand("select num_sch from number_sch where name_rp = '" + form3.comboBox1.SelectedItem.ToString() + "' and name_tii = '" + form3.checkedListBox1.Items[i].ToString() + "'", connection);
                    //OleDbDataReader datar = com.ExecuteReader();
                    NpgsqlCommand com = new NpgsqlCommand("select \"num_sch\" from \"number_sch\" where \"name_rp\" = '" + options.comboBox1.SelectedItem.ToString() + "' and \"name_tii\" = '" + options.checkedListBox1.Items[i].ToString() + "'", connection);
                    NpgsqlDataReader datar = com.ExecuteReader();

                    List<string> list = new List<string>();

                    while (datar.Read())
                    {
                        list.Add(datar.GetValue(0).ToString());
                        //list.Add(datar.GetValue(1).ToString());
                        //list.Add(datar.GetValue(2).ToString());
                    }
                    this.dataGridView2.Rows[i].Cells[1].Value = list[0];
                    //this.dataGridView2.Rows[i].Cells[1].Value = list[1];
                    //this.dataGridView2.Rows[i].Cells[3].Value = list[2];

                }
                connection.Close();
                this.dataGridView2.Rows[i].Cells[0].Value = options.checkedListBox1.Items[i].ToString();
                if (this.dataGridView2.Rows[i].Cells[1].Value == null) { this.dataGridView2.Rows[i].Cells[1].Value = 0; }

            }
            string str = "";
            ls.Clear();
            dataGridView3.DataSource = null;
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                str = dataGridView2.Rows[i].Cells[0].Value.ToString();
                if (str.Substring(0, 3).ToString() == "пок" || str.Substring(0, 3).ToString() == "Пок")
                {
                    mails m = new mails();
                    m.name_tii = dataGridView2.Rows[i].Cells[0].Value.ToString();
                    m.num_sch = dataGridView2.Rows[i].Cells[1].Value.ToString();
                    ls.Add(m);
                }
            }
            dataGridView3.DataSource = ls;
            dataGridView3.Columns[0].HeaderText = "Присоединение";
            dataGridView3.Columns[1].HeaderText = "Счетчик №";
            dataGridView3.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells); //автоматический размер колонок
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DateTime data = DateTime.Now;  //Время запуска программы
            String host = System.Net.Dns.GetHostName(); //
            System.Net.IPAddress ip = System.Net.Dns.GetHostByName(host).AddressList[0];


            NpgsqlConnection connection = new NpgsqlConnection(put);
            connection.Open();
            for (int i = 0; i < this.dataGridView3.Rows.Count; i++)
            {
                if (this.dataGridView3.Rows[i].Cells[1].Value == null) { this.dataGridView3.Rows[i].Cells[1].Value = 0; }
                //OleDbCommand command1 = new OleDbCommand("select count(*) from number_sch where name_rp = '" + form3.comboBox1.SelectedItem.ToString() + "' and name_tii = '" + this.dataGridView3.Rows[i].Cells[0].Value.ToString() + "'", connection);
                NpgsqlCommand command1 = new NpgsqlCommand("select count(*) from \"number_sch\" where \"name_rp\" = '" + options.comboBox1.SelectedItem.ToString() + "' and \"name_tii\" = '" + this.dataGridView3.Rows[i].Cells[0].Value.ToString() + "'", connection);

                int p = Convert.ToInt32(command1.ExecuteScalar());

                NpgsqlCommand command6 = new NpgsqlCommand("select count(*) from \"journal\" where \"name_rp\" = '" + options.comboBox1.SelectedItem.ToString() + "' and \"name_tii\" = '" + this.dataGridView3.Rows[i].Cells[0].Value.ToString() + "'", connection);
                int p1 = Convert.ToInt32(command6.ExecuteScalar());

                if (p == 0)
                {

                    NpgsqlCommand command = new NpgsqlCommand("insert into \"number_sch\" values('" + options.comboBox1.SelectedItem.ToString() + "', '" + this.dataGridView3.Rows[i].Cells[0].Value.ToString() + "', '" + this.dataGridView3.Rows[i].Cells[1].Value.ToString() + "')", connection);
                    command.ExecuteNonQuery();

                }

                if (p1 == 0)
                {
                    NpgsqlCommand command4 = new NpgsqlCommand("insert into \"journal\" values('" + data.ToString() + "', '" + ip.ToString() + "', '" + options.comboBox1.SelectedItem.ToString() + "', '" + this.dataGridView3.Rows[i].Cells[0].Value.ToString() + "', 'nomer' , 'schet',  '" + this.dataGridView3.Rows[i].Cells[1].Value.ToString() + "')", connection);
                    command4.ExecuteNonQuery();
                }
                if (p1 != 0)
                {
                    NpgsqlCommand command4 = new NpgsqlCommand("insert into \"journal\" values('" + data.ToString() + "', '" + ip.ToString() + "', '" + options.comboBox1.SelectedItem.ToString() + "', '" + this.dataGridView3.Rows[i].Cells[0].Value.ToString() + "', 'nomer' , 'schet',  '" + this.dataGridView3.Rows[i].Cells[1].Value.ToString() + "')", connection);
                    command4.ExecuteNonQuery();
                }

                if (p == 1)
                {
                    NpgsqlCommand command3 = new NpgsqlCommand("update \"number_sch\" set \"num_sch\" = '" + this.dataGridView3.Rows[i].Cells[1].Value.ToString() + "' where \"name_rp\" = '" + options.comboBox1.SelectedItem.ToString() + "' and \"name_tii\" = '" + this.dataGridView3.Rows[i].Cells[0].Value.ToString() + "'", connection);
                    command3.ExecuteNonQuery();

                }
            }
        }

        private void Options_Load(object sender, EventArgs e)
        {
            String host = System.Net.Dns.GetHostName(); //
            System.Net.IPAddress ip = System.Net.Dns.GetHostByName(host).AddressList[0];

            comboBox1.Items.Clear();

            NpgsqlConnection connection = new NpgsqlConnection(put);
            connection.Open();
            NpgsqlCommand command1 = new NpgsqlCommand("select count(*) from \"users_admins\" where \"ip\" = '" + ip.ToString() + "'", connection);
            int p = Convert.ToInt32(command1.ExecuteScalar());

                    if (p == 1)
                    {
                                status = "Администратор";
                                label3.Text = "Режим редактирования.";
                                label3.ForeColor = Color.Green;
                                button2.Enabled = true;
                                button4.Enabled = true;
                    }
                            if (p==0)
                            {
                                status = "Пользователь";
                                this.dataGridView1.Columns[0].ReadOnly = true;
                                this.dataGridView1.Columns[1].ReadOnly = true;
                                this.dataGridView1.Columns[2].ReadOnly = true;
                                this.dataGridView1.Columns[3].ReadOnly = true;
                                this.dataGridView1.Columns[5].ReadOnly = true;

                                label3.Text = "Режим чтения. Редактирование ЗАПРЕЩЕНО!";
                                label3.ForeColor = Color.Red;

                            }

                            comboBox1.Items.Clear();
                            NpgsqlCommand command2 = new NpgsqlCommand("select client_addr from pg_stat_activity where datname = 'monitor'", connection);
                       NpgsqlDataReader datar = command2.ExecuteReader();

                    while (datar.Read())
                    {
                        comboBox1.Items.Add(datar.GetValue(0).ToString());
                    }

                connection.Close();
                this.Text = options.comboBox1.SelectedItem.ToString() + "  " + host + "  " + status;
                if (ip.ToString() == "192.168.145.87") { button5.Enabled = true; }

            dataGridView1.Rows.Clear();
            dataGridView3.DataSource = null;
            button1_Click(sender,e);
            button3_Click(sender, e);

        }

        private void button5_Click(object sender, EventArgs e)
        {
            NpgsqlConnection connection = new NpgsqlConnection(put);
            connection.Open();
            string zapros = "select pg_terminate_backend(pid), datname from pg_stat_activity where datname = 'monitor'";
            NpgsqlCommand command1 = new NpgsqlCommand(zapros, connection);
            command1.Prepare();
            command1.ExecuteNonQuery();
            connection.Close();


        }

    }
}
