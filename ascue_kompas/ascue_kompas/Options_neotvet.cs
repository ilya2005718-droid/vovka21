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
//using Npgsql;
using Npgsql;
using ExcelLibrary;
using ExcelLibrary.SpreadSheet;
using ExcelLibrary.CompoundDocumentFormat;
using ExcelLibrary.BinaryFileFormat;


namespace ascue_kompas
{
    public partial class Options_neotvet : Form
    {
        private View options_neotvet;
        
        public Options_neotvet(View on)
        {
            InitializeComponent();
            options_neotvet = on;
        }

        class class_new
        {
            public string id { get; set; }
            public string name_rp { get; set; }
            public string name_tii { get; set; }
        }
        class class_old
        {
            public string id { get; set; }
            public string name_rp { get; set; }
            public string name_tii { get; set; }
        }

        List<class_new> list_new = new List<class_new>();
        List<class_old> list_old = new List<class_old>();


        public string monitor = "Server=192.168.150.40;Port=5432;;User Id=psql; Password=;Database=monitor; CommandTimeout=955555";
        public string monitor_options = "Server=192.168.150.40;Port=5432;;User Id=psql; Password=;Database=monitor_options; CommandTimeout=955555";
        
        public DataGridView zagruzka_new_options_neotvet(DataGridView dgv)
        {
            string z = "";
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            List<string> list_rp = new List<string>();
            NpgsqlConnection connection = new NpgsqlConnection(monitor);
            DataTable table = new DataTable();
            connection.Open();

            NpgsqlCommand com_rp = new NpgsqlCommand("select \"Name\" from \"GROUPS\" where \"Parent\" is null order by regexp_replace(\"Name\", '[^0-9]', '', 'g')::numeric ", connection);
            NpgsqlDataReader datar = com_rp.ExecuteReader();
            while (datar.Read()) { list_rp.Add(datar.GetString(0)); }
            dgv.Rows.Clear();
            for (int t = 0; t < list_rp.Count; t++)
            {
                z = "select t4.\"Values_ID\", t0.\"Name\", t4.\"Name\" from (select t1.\"Groups_ID\", t2.\"Name\" from \"GROUPS\" t1, \"GROUPS\" t2 where t1.\"Tag\" = 'ТИИ' and t1.\"Parent\"=t2.\"Groups_ID\" and t2.\"Name\" = '" + list_rp[t].ToString() + "') as t0, \"TB_VALUES\" t4 where t4.\"Groups_ID\"=t0.\"Groups_ID\" and lower(t4.\"Name\") like 'пок%'";
                NpgsqlCommand com = new NpgsqlCommand(z, connection);
                NpgsqlDataAdapter dap = new NpgsqlDataAdapter(com);
                dap.Fill(table);
                connection.Close();
                list_new.Clear();
                foreach (DataRow row in table.Rows)
                {
                    class_new m = new class_new();
                    m.id = row.ItemArray[0].ToString();
                    m.name_rp = row.ItemArray[1].ToString();
                    m.name_tii = row.ItemArray[2].ToString();
                    list_new.Add(m);
                }
            }

            NpgsqlConnection connection1 = new NpgsqlConnection(monitor_options);
            DataTable table1 = new DataTable();
            connection1.Open();
            z = "select \"id\",\"name_rp\",\"name_tii\" from neotvet order by regexp_replace(\"name_rp\", '[^0-9]', '', 'g')::numeric";
            NpgsqlCommand com1 = new NpgsqlCommand(z, connection1);
            NpgsqlDataAdapter dap1 = new NpgsqlDataAdapter(com1);
            dap1.Fill(table1);
            connection1.Close();
            list_old.Clear();
            foreach (DataRow row in table1.Rows)
            {
                class_old m = new class_old();
                m.id = row.ItemArray[0].ToString();
                m.name_rp = row.ItemArray[1].ToString();
                m.name_tii = row.ItemArray[2].ToString();
                list_old.Add(m);
            }
            //MessageBox.Show(list_new.Count.ToString() + "   " + list_old.Count.ToString());


            foreach (var item in list_new)
            {
                if (list_old.Where(c => c.id == item.id).Count() == 0)
                {
                    int rowNum = dgv.Rows.Add();
                    
                    dgv.Rows[rowNum].Cells[0].Value = item.id.ToString();
                    dgv.Rows[rowNum].Cells[1].Value = item.name_rp.ToString();
                    dgv.Rows[rowNum].Cells[2].Value = item.name_tii.ToString();
                    dgv.Rows[rowNum].Cells[3].Value = false;
                    dgv.Rows[rowNum].Cells[4].Value = false;
                    dgv.Rows[rowNum].Cells[5].Value = false;
                    
                }
            }


            MessageBox.Show("Всего " + dgv.Rows.Count.ToString() + " новых присоединений");

            connection.Close();
            
                return dgv;
            
        }
        public DataGridView zagruzka_options_neotvet(DataGridView dgv)
        {
            string z = "";
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgv.Rows.Clear();
            NpgsqlConnection connection = new NpgsqlConnection(monitor_options);
            connection.Open();
            z = "select \"id\",\"name_rp\",\"name_tii\",\"vis\",\"niz\",\"rezerv\" from neotvet order by regexp_replace(\"name_rp\", '[^0-9]', '', 'g')::numeric";
            NpgsqlCommand com = new NpgsqlCommand(z, connection);
                NpgsqlDataReader reader = com.ExecuteReader();
                List<string[]> data = new List<string[]>();

                while (reader.Read())
                {
                    data.Add(new string[6]);

                    data[data.Count - 1][0] = reader[0].ToString();
                    data[data.Count - 1][1] = reader[1].ToString();
                    data[data.Count - 1][2] = reader[2].ToString();
                    data[data.Count - 1][3] = reader[3].ToString();
                    data[data.Count - 1][4] = reader[4].ToString();
                    data[data.Count - 1][5] = reader[5].ToString();
                }

                reader.Close();
                
                foreach (string[] s in data)
                    dgv.Rows.Add(s);
               
            connection.Close();
            MessageBox.Show("Всего " + dgv.Rows.Count.ToString() + " присоединений");
                return dgv;
        }
        public void save_options_neotvet(DataGridView dgv)
        { 
            NpgsqlConnection connection = new NpgsqlConnection(monitor_options);
            connection.Open();
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                NpgsqlCommand command1 = new NpgsqlCommand("select count(*) from \"neotvet\" where \"id\" = '" + dgv.Rows[i].Cells[0].Value.ToString() + "'", connection);
                int p = Convert.ToInt32(command1.ExecuteScalar());
 
                
                    if (p == 0)
                    {
                        
                        string z = "insert into \"neotvet\" values('" + dgv.Rows[i].Cells[0].Value.ToString() + "', '" + dgv.Rows[i].Cells[1].Value.ToString() + "', '" + dgv.Rows[i].Cells[2].Value.ToString() + "', " + dgv.Rows[i].Cells[3].Value.ToString() + ", " + dgv.Rows[i].Cells[4].Value.ToString() + ", " + dgv.Rows[i].Cells[5].Value.ToString() + ")";
                        NpgsqlCommand command = new NpgsqlCommand(z, connection);
                        command.ExecuteNonQuery();
                    }
                    if (p == 1)
                    {
                        NpgsqlCommand command3 = new NpgsqlCommand("update \"neotvet\" set \"id\" = '" + dgv.Rows[i].Cells[0].Value.ToString() + "', \"name_rp\" = '" + dgv.Rows[i].Cells[1].Value.ToString() + "', \"name_tii\" = '" + dgv.Rows[i].Cells[2].Value.ToString() + "', \"vis\" = " + dgv.Rows[i].Cells[3].Value.ToString() + ", \"niz\" = " + dgv.Rows[i].Cells[4].Value.ToString() + ", \"rezerv\" = " + dgv.Rows[i].Cells[5].Value.ToString() + " where \"id\" = '" + dgv.Rows[i].Cells[0].Value.ToString() + "'", connection);
                        command3.ExecuteNonQuery();
                    }
                
            }
            connection.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            zagruzka_new_options_neotvet(this.dataGridView1);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            save_options_neotvet(this.dataGridView1);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            zagruzka_options_neotvet(this.dataGridView1);
        }

        private void Options_neotvet_Load(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
        }

    }
}
