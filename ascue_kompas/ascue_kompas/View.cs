using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Windows.Forms.DataVisualization.Charting;
using ascue_kompas;
using System.IO;
using Npgsql;
using ExcelLibrary;
using ExcelLibrary.SpreadSheet;
using ExcelLibrary.CompoundDocumentFormat;
using ExcelLibrary.BinaryFileFormat;
namespace ascue_kompas
{
    public partial class View : Form
    {
        Controller cn;
        Options options;
        Options_neotvet options_neotvet;
        public void setController(Controller cont)
        {
            cn = cont;
        }
        public View()
        {
            CheckForIllegalCrossThreadCalls = false;
            InitializeComponent();
            options = new Options(this);
            options_neotvet = new Options_neotvet(this);
        }
        public string put = "Server=192.168.150.40;Port=5432;;User Id=psql; Password=;Database=monitor_options; CommandTimeout=955555";
        private void View_Load(object sender, EventArgs e)
        {

            String host = System.Net.Dns.GetHostName(); 
            System.Net.IPAddress ip = System.Net.Dns.GetHostByName(host).AddressList[0];
            string status = "";
            cb_base.SelectedItem = "Компас РП";
            NpgsqlConnection connection = new NpgsqlConnection(put);
            connection.Open();
            NpgsqlCommand command1 = new NpgsqlCommand("select count(*) from \"users_admins\" where \"ip\" = '" + ip.ToString() + "'", connection);
            int p = Convert.ToInt32(command1.ExecuteScalar());
            if (p == 1)
            {
               status = "Администратор";
            }
                 
            if (p == 0)
            {
                status = "Пользователь";
            }
            connection.Close();
            
            this.Text = this.Text + "   " + host + "   " + ip.ToString() + "   "+status;

            cn.logy();
            cn.zagruzka_spiska_rp();
            cn.zagruzka_spiska_prisoed();
            contextMenuStrip1.Items[4].Enabled = false;
            contextMenuStrip1.Items[6].Enabled = false;
            contextMenuStrip1.Items[7].Enabled = false;
            dateTimePicker1.Value = DateTime.Now;
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            if (cb_base.SelectedItem.ToString() == "Компас РП" || cb_base.SelectedItem.ToString() == "Компас РП измер")
            {
                cn.zagruzka_spiska_othod_rp();
                cn.zagruzka_spiska_prisoed();
            }

           if (cb_base.SelectedItem.ToString() == "Энергомера")
            {
                cn.zagruzka_spiska_othod_rp_energo();
                cn.zagruzka_spiska_prisoed();
            }
            if (cb_base.SelectedItem.ToString() == "Энергосфера")
            {
                cn.zagruzka_spiska_othod_r();
            }

            if (cb_base.SelectedItem.ToString() == "Меркурии")
            {
                cn.zagruzka_spiska_oth_tp_mercury();
            }

            if (cb_base.SelectedItem.ToString() == "Энергомера ТП")
            {
                cn.zagruzka_spiska_othod_tp_energo();
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            cn.zagolovok = "Данные счетчиков " + comboBox1.SelectedItem.ToString();
            label7.Text = cn.zagolovok;
            pictureBox1.Visible = true;
            Thread t = new Thread(new ThreadStart(delegate
            {
                    this.Invoke(new ThreadStart(delegate
                    {
                        try
                        {
                            if (cb_base.SelectedItem.ToString() == "Компас РП" || cb_base.SelectedItem.ToString() == "Компас РП измер")
                            {
                                cn.zagruzka_resultat();
                            }
                            if (cb_base.SelectedItem.ToString() == "Энергосфера")
                            {
                                cn.zagruzka_resultat_sphera();
                            }
                            if (cb_base.SelectedItem.ToString() == "Энергомера") 
                            {
                                cn.zagruzga_resultat_energomera(); 
                            }
                            if (cb_base.SelectedItem.ToString() == "Меркурии")
                            {
                                cn.zagruzka_resultat_mercury();
                            }
                            if (cb_base.SelectedItem.ToString() == "Энергомера ТП")
                            {
                                cn.zagruzka_resultat_tp_energo();
                            }

                            pictureBox1.Visible = false;
                        }
                        catch { }

                        }));
                }));
            t.Start();
            развернутьПрисоединенияToolStripMenuItem1.Enabled = false; 
        }
        private void графикПоДнямToolStripMenuItem_Click(object sender, EventArgs e)
        {
            cn.zagruzka_graphic();
            comboBox7.Enabled = true;
        }
        private void таблицаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            chart1.Visible = false;
            chart2.Visible = false;
            chart3.Visible = false;
            comboBox7.Enabled = false;
            comboBox6.Enabled = false;
        }
        private void dataGridView3_MouseUp(object sender, MouseEventArgs e)
        {
            if (checkBox4.Checked == true)
            {
                contextMenuStrip1.Items[0].Enabled = true;
                if (e.Button == MouseButtons.Right)
                { contextMenuStrip1.Show(MousePosition, ToolStripDropDownDirection.Right); }
            }
            if (checkBox4.Checked == false) { contextMenuStrip1.Items[0].Enabled = false; }
        }
        private void chart1_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            { contextMenuStrip2.Show(MousePosition, ToolStripDropDownDirection.Right); }
        }
        private void checkBox1_Click(object sender, EventArgs e)
        {
            cn.check_pok();
        }
        private void checkBox2_Click(object sender, EventArgs e)
        {
            cn.check_sutki();
        }
        private void checkBox3_Click(object sender, EventArgs e)
        {
            cn.check_p();
        }
        private void небалансРУ610кВToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DateTime d1 = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day); //дата начала периода
            DateTime d2 = new DateTime(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day); //дата конца периода
            cn.zagolovok = "Информация о небалансе  " + comboBox1.SelectedItem.ToString() + "   за период   " + d1 + "   -   " + d2 + " РУ-6-10 кВ";
            label7.Text = cn.zagolovok;

            pictureBox1.Visible = true;
            Thread t = new Thread(new ThreadStart(delegate
            {
                this.Invoke(new ThreadStart(delegate
                {
                    cn.nebalans_v();
                    pictureBox1.Visible = false;

                }));
            }));
            t.Start();

        }
        private void chart2_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            { contextMenuStrip1.Show(MousePosition, ToolStripDropDownDirection.Right); }
        }
        private void графикНебалансаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            chart2.Visible = true;
            chart2.BringToFront();
            comboBox6.Enabled = true;
            comboBox7.Enabled = true;
            comboBox8.Visible = false;
            comboBox6.Visible = true;
        }
        private void chart2_MouseUp_1(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            { contextMenuStrip3.Show(MousePosition, ToolStripDropDownDirection.Right); }
        }
        private void небалансРУ04КВToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DateTime d1 = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day); //дата начала периода
            DateTime d2 = new DateTime(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day); //дата конца периода
            cn.zagolovok = "Информация о небалансе  " + comboBox1.SelectedItem.ToString() + "   за период   " + d1 + "   -   " + d2 + " РУ-0,4 кВ";
            label7.Text = cn.zagolovok;

            pictureBox1.Visible = true;
            Thread t = new Thread(new ThreadStart(delegate
            {
                this.Invoke(new ThreadStart(delegate
                {
                    cn.nebalans_n();
                    pictureBox1.Visible = false;

                }));
            }));
            t.Start();


        }
        private void суточноеПотреблениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DateTime d1 = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day); //дата начала периода
            DateTime d2 = new DateTime(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day); //дата конца периода
            cn.zagolovok = "Информация о суточном потреблении  " + comboBox1.SelectedItem.ToString() + "   за период   " + d1 + "   -   " + d2;
            label7.Text = cn.zagolovok;
            pictureBox1.Visible = true;
            Thread t = new Thread(new ThreadStart(delegate
            {
                this.Invoke(new ThreadStart(delegate
                {
                    cn.sutochnoe_potreb();
                    pictureBox1.Visible = false;

                }));
            }));
            t.Start();


        }
        private void графикСуточПотребToolStripMenuItem_Click(object sender, EventArgs e)
        {
            chart2.Visible = true;
            chart2.BringToFront();
            comboBox7.Enabled = true;
        }
        private void button6_Click(object sender, EventArgs e)
        {
            DateTime d1 = new DateTime(dateTimePicker3.Value.Year, dateTimePicker3.Value.Month, dateTimePicker3.Value.Day);
            if (checkBox5.Checked == true)
            {
                cn.zagolovok = "Информация о суточном небалансе  " + comboBox1.SelectedItem.ToString() + "   дата  " + d1 + "  РУ - 0,4 кВ";
            }
            else
            {
                cn.zagolovok = "Информация о суточном небалансе  " + comboBox1.SelectedItem.ToString() + "   дата  " + d1 + "  РУ - 6-10 кВ";
            }
            label7.Text = cn.zagolovok;

            pictureBox1.Visible = true;
            Thread t = new Thread(new ThreadStart(delegate
            {
                this.Invoke(new ThreadStart(delegate
                {
                    if (cb_base.SelectedItem.ToString() == "Компас РП")
                    {
                        cn.sutochniy_nebalans();
                    }
                    if (cb_base.SelectedItem.ToString() == "Энергомера") { cn.sutochniy_nebalans_energomera(); }
                    pictureBox1.Visible = false;

                }));
            }));
            t.Start();
            
            развернутьПрисоединенияToolStripMenuItem1.Enabled = true;
        }
        private void графикРПСуткиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            chart3.Visible = true;
            chart3.BringToFront();
            comboBox7.Enabled = true;
            comboBox6.Enabled = true;
            comboBox8.Visible = true;
            comboBox8.Enabled = true;
            comboBox6.Visible = false;
        }
        private void chart3_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            { contextMenuStrip4.Show(MousePosition, ToolStripDropDownDirection.Right); }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            DateTime d1 = new DateTime(dateTimePicker3.Value.Year, dateTimePicker3.Value.Month, dateTimePicker3.Value.Day);
            cn.zagolovok = "Информация о суточном небалансе  " + comboBox1.SelectedItem.ToString() + "   дата  " + d1 + "РУ-6-10 кВ";
            label7.Text = cn.zagolovok;

            pictureBox1.Visible = true;
            Thread t = new Thread(new ThreadStart(delegate
            {
                this.Invoke(new ThreadStart(delegate
                {
                    if (cb_base.SelectedItem.ToString() == "Компас РП")
                    {
                        cn.sutochniy_nebalans_prisoed();
                    }
                    if (cb_base.SelectedItem.ToString() == "Энергомера")
                    {
                        cn.sutochniy_nebalans_prisoed_energo();
                    }
                    pictureBox1.Visible = false;

                }));
            }));
            t.Start();

            развернутьПрисоединенияToolStripMenuItem1.Enabled = true;

        }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            DateTime d1 = new DateTime(dateTimePicker3.Value.Year, dateTimePicker3.Value.Month, dateTimePicker3.Value.Day);
            cn.zagolovok = "Информация о суточном небалансе  " + comboBox1.SelectedItem.ToString() + "   дата  " + d1;
            label7.Text = cn.zagolovok;
            pictureBox1.Visible = true;
            Thread t = new Thread(new ThreadStart(delegate
                {
                    this.Invoke(new ThreadStart(delegate
                        {
                            if (cb_base.SelectedItem.ToString() == "Компас РП")
                            {
                                cn.sutochniy_nebalans_prisoed();
                                cn.sut_nebalans_povremeni();
                            }
                            if (cb_base.SelectedItem.ToString() == "Энергомера")
                            {
                                cn.sutochniy_nebalans_prisoed_energo();
                                cn.sut_nebalans_povremeni();
                            }
                            pictureBox1.Visible = false;
                        }));
                }));
            t.Start();
        }
        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            DateTime d1 = new DateTime(dateTimePicker3.Value.Year, dateTimePicker3.Value.Month, dateTimePicker3.Value.Day);
            cn.zagolovok = "Информация о суточном потреблении  " + comboBox5.SelectedItem.ToString() + "   дата  " + d1;
            label7.Text = cn.zagolovok;

            pictureBox1.Visible = true;
            Thread t = new Thread(new ThreadStart(delegate
                {
                    this.Invoke(new ThreadStart(delegate
                        {
                            if (cb_base.SelectedItem.ToString() == "Компас РП")
                            {
                                cn.sutochniy_nebalans_prisoed();
                                cn.sutochniy_nebalans_prisoed_poothod();
                            }
                            if (cb_base.SelectedItem.ToString() == "Энергомера")
                            {
                                cn.sutochniy_nebalans_prisoed_energo();
                                cn.sutochniy_nebalans_prisoed_poothod();
                            }
                            pictureBox1.Visible = false;

                        }));
                }));
            t.Start();
        }
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            cn.print_page(e);
        }
        private void printDocument1_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {
            cn.print_begin();
        }
        private void таблицаToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            chart1.Visible = false;
            chart2.Visible = false;
            chart3.Visible = false;
            comboBox7.Enabled = false;
            comboBox6.Enabled = false;
        }
        private void печатьГрафикаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            chart1.Printing.Print(true);
        }
        private void печатьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            chart2.Printing.Print(true);
        }
        private void печатьToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            chart3.Printing.Print(true);
        }
        private void таблицаToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            chart1.Visible = false;
            chart2.Visible = false;
            chart3.Visible = false;
            comboBox7.Enabled = false;
            comboBox6.Enabled = false;
            comboBox8.Enabled = false;
        }
        private void таблицаToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            chart1.Visible = false;
            chart2.Visible = false;
            chart3.Visible = false;
            comboBox7.Enabled = false;
            comboBox6.Enabled = false;
        }
        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            cn.vid_graphica();
        }
        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            cn.vid_graphica_proc();
        }
        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {
            cn.vid_graphica_proc_sutki();
        }
        private void сохранитьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            cn.export_to_excel();
        }
        private void печатьToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            //Open the print dialog
            PrintDialog printDialog = new PrintDialog();
            printDialog.Document = printDocument1;
            printDialog.UseEXDialog = true;

            //Get the document
            if (DialogResult.OK == printDialog.ShowDialog())
            {
                printDocument1.DocumentName = "Test Page Print";
                printDocument1.Print();
            }
        }
        private void выходToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private void настройкиToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            options.ShowDialog();
        }
        private void развернутьПрисоединенияToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            cn.povernut_prisoed();
        }
        private void инструкцияToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Form instruk = new Form();
            instruk.Text = "Инструкция";
            instruk.Width = 680;
            instruk.Height = 600;
            RichTextBox txtb = new RichTextBox();
            txtb.Text = File.ReadAllText("instructions.txt", Encoding.Default);
            txtb.Visible = true;
            txtb.Width = 650;
            txtb.Height = 550;
            txtb.Location = new Point(5, 5);
            instruk.Controls.Add(txtb);
            instruk.Show();
        }
        Form j_zaprosov = new Form();
        private void журналЗапросовToolStripMenuItem_Click(object sender, EventArgs e)
        {
            j_zaprosov.Text = "Журнал запросов";
            j_zaprosov.Width = 680;
            j_zaprosov.Height = 600;
            richTextBox1.Visible = true;
            richTextBox1.Width = 650;
            richTextBox1.Height = 500;
            richTextBox1.Location = new Point(5, 5);
            Button zakr = new Button();
            zakr.Width = 100;
            zakr.Height = 30;
            zakr.Text = "Закрыть";
            zakr.Location = new Point(20, 520);
            zakr.Click += new EventHandler(zakr_Click);
            j_zaprosov.Controls.Add(richTextBox1);
            j_zaprosov.Controls.Add(zakr);
            j_zaprosov.Show();
        }
        void zakr_Click(object sender, EventArgs e)
        {
            j_zaprosov.Hide();
        }
        private void небалансРППериодToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DateTime d1 = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day); //дата начала периода
            DateTime d2 = new DateTime(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day); //дата конца периода
            cn.zagolovok = "Информация о небалансе  " + comboBox1.SelectedItem.ToString() + "   за период   " + d1 + "   -   " + d2;
            label7.Text = cn.zagolovok;

            pictureBox1.Visible = true;
            Thread t = new Thread(new ThreadStart(delegate
            {
                this.Invoke(new ThreadStart(delegate
                {
                    cn.nebalans_v_per();
                    pictureBox1.Visible = false;

                }));
            }));
            t.Start();
        }
        private void замечанияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form notes = new Form();
            notes.Text = "Замечания";
            notes.Width = 680;
            notes.Height = 600;
            RichTextBox txtb = new RichTextBox();
            txtb.Text = File.ReadAllText("notes.txt", Encoding.Default);
            txtb.Visible = true;
            txtb.Width = 650;
            txtb.Height = 550;
            txtb.Location = new Point(5, 5);
            notes.Controls.Add(txtb);
            notes.Show();
        }
        private void настройкиНеответToolStripMenuItem_Click(object sender, EventArgs e)
        {
            options_neotvet.ShowDialog();
        }
        private void ру610КВToolStripMenuItem_Click(object sender, EventArgs e)
        {
            cn.zagolovok = "Недостоверные данные счетчиков Ру-6-10 кВ";
            label7.Text = cn.zagolovok;
            pictureBox1.Visible = true;
            Thread t = new Thread(new ThreadStart(delegate
            {

                this.Invoke(new ThreadStart(delegate
                {
                    try
                    {
                        cn.report_neotvet_vis();
                        pictureBox1.Visible = false;
                    }
                    catch { }

                }));
            }));
            t.Start();
            
        }
        private void ру04КВToolStripMenuItem_Click(object sender, EventArgs e)
        {
            cn.zagolovok = "Недостоверные данные счетчиков Ру-0,4 кВ";
            label7.Text = cn.zagolovok;
            pictureBox1.Visible = true;
            Thread t = new Thread(new ThreadStart(delegate
            {

                this.Invoke(new ThreadStart(delegate
                {
                    try
                    {
                        cn.report_neotvet_niz();
                        pictureBox1.Visible = false;
                    }
                    catch { }

                }));
            }));
            t.Start();
            
        }
        private void резервToolStripMenuItem_Click(object sender, EventArgs e)
        {
            cn.zagolovok = "Недостоверные данные счетчиков в резерве";
            label7.Text = cn.zagolovok;
            pictureBox1.Visible = true;
            Thread t = new Thread(new ThreadStart(delegate
            {

                this.Invoke(new ThreadStart(delegate
                {
                    try
                    {
                        cn.report_neotvet_rezerv();
                        pictureBox1.Visible = false;
                    }
                    catch { }

                }));
            }));
            t.Start();
            
        }
        private void всеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            cn.zagolovok = "Недостоверные данные всех счетчиков ";
            label7.Text = cn.zagolovok;
            pictureBox1.Visible = true;
            Thread t = new Thread(new ThreadStart(delegate
            {

                this.Invoke(new ThreadStart(delegate
                {
                    try
                    {
                        cn.report_neotvet_all();
                        pictureBox1.Visible = false;
                    }
                    catch { }

                }));
            }));
            t.Start();
            
        }
        private void cb_base_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (cb_base.SelectedItem.ToString() == "Компас РП")
            {
                cn.zagruzka_spiska_rp();
                checkBox1.Enabled = true;
                checkBox3.Enabled = true;
                checkBox2.Enabled = false;
                checkBox4.Enabled = true;

                checkBox_v.Checked = false;
                checkBox_z.Checked = false;
                checkBox_s.Checked = false;
                checkBox_u.Checked = false;

                checkBox_v.Enabled = false;
                checkBox_s.Enabled = false;
                checkBox_z.Enabled = false;
                checkBox_u.Enabled = false;

                textBox_poisk.Enabled = false;

                dateTimePicker3.Enabled = true;
                button5.Enabled = true;
                button6.Enabled = true;
                comboBox2.Enabled = true;
                comboBox5.Enabled = true;
                checkBox5.Enabled = true;
                contextMenuStrip1.Enabled = true;
                
            }
           
            
            if (cb_base.SelectedItem.ToString() == "Компас РП измер")
            {
                cn.zagruzka_spiska_rp();
                checkBox1.Enabled = false;
                checkBox3.Enabled = false;
                checkBox2.Enabled = false;
                checkBox4.Enabled = false;
                
                checkBox_v.Checked = false;
                checkBox_z.Checked = false;
                checkBox_s.Checked = false;
                checkBox_u.Checked = false;

                checkBox_v.Enabled = false;
                checkBox_s.Enabled = false;
                checkBox_z.Enabled = false;
                checkBox_u.Enabled = false;

                textBox_poisk.Enabled = false;

                dateTimePicker3.Enabled = false;
                button5.Enabled = false;
                button6.Enabled = false;
                comboBox2.Enabled = false;
                comboBox5.Enabled = false;
                checkBox5.Enabled = false;
                contextMenuStrip1.Enabled = false;
               
            }
            
            if (cb_base.SelectedItem.ToString() == "Энергомера")
            {
                cn.zagruzka_spiska_rp_energo();
                checkBox1.Enabled = false;
                checkBox3.Enabled = false;
                checkBox2.Enabled = false;
                checkBox4.Enabled = true;
                
                checkBox_v.Checked = false;
                checkBox_z.Checked = false;
                checkBox_s.Checked = false;
                checkBox_u.Checked = false;

                checkBox_v.Enabled = false;
                checkBox_s.Enabled = false;
                checkBox_z.Enabled = false;
                checkBox_u.Enabled = false;

                textBox_poisk.Enabled = false;

                dateTimePicker3.Enabled = true;
                button5.Enabled = true;
                button6.Enabled = true;
                comboBox2.Enabled = true;
                comboBox5.Enabled = true;
                checkBox5.Enabled = true;
                contextMenuStrip1.Enabled = true;
            }


            
            if (cb_base.SelectedItem.ToString() == "Энергосфера")
            {
                cn.zagruzka_spiska_r();
                checkBox1.Enabled = false;
                checkBox3.Enabled = false;
                checkBox2.Enabled = false;
                checkBox4.Enabled = false;
                
                checkBox_v.Checked = false;
                checkBox_z.Checked = false;
                checkBox_s.Checked = false;
                checkBox_u.Checked = false;

                checkBox_v.Enabled = false;
                checkBox_s.Enabled = false;
                checkBox_z.Enabled = false;
                checkBox_u.Enabled = false;

                textBox_poisk.Enabled = false;

                dateTimePicker3.Enabled = false;
                button5.Enabled = false;
                button6.Enabled = false;
                comboBox2.Enabled = false;
                comboBox5.Enabled = false;
                checkBox5.Enabled = false;
                contextMenuStrip1.Enabled = false;
            }

            if (cb_base.SelectedItem.ToString() == "Меркурии")
            {
                cn.zagruzka_spiska_tp_mercury();
                checkBox1.Enabled = false;
                checkBox3.Enabled = false;
                checkBox2.Enabled = true;
                checkBox4.Enabled = false;
                
                checkBox_v.Checked = false;
                checkBox_z.Checked = false;
                checkBox_s.Checked = false;
                checkBox_u.Checked = false;

                checkBox_v.Enabled = false;
                checkBox_s.Enabled = false;
                checkBox_z.Enabled = false;
                checkBox_u.Enabled = false;

                textBox_poisk.Enabled = false;
                
                dateTimePicker3.Enabled = false;
                button5.Enabled = false;
                button6.Enabled = false;
                comboBox2.Enabled = false;
                comboBox5.Enabled = false;
                checkBox5.Enabled = false;
                contextMenuStrip1.Enabled = false;
            }
            if (cb_base.SelectedItem.ToString() == "Энергомера ТП")
            {
                
                checkBox1.Enabled = false;
                checkBox3.Enabled = false;
                checkBox2.Enabled = false;
                checkBox4.Enabled = false;
                checkBox_v.Enabled = true;
                checkBox_s.Enabled = true;
                checkBox_z.Enabled = true;
                checkBox_u.Enabled = true;
                checkBox_v.Checked = true;
                checkBox_z.Checked = false;
                checkBox_s.Checked = false;
                checkBox_u.Checked = false;
                
                textBox_poisk.Enabled = true;
                
                dateTimePicker3.Enabled = false;
                button5.Enabled = false;
                button6.Enabled = false;
                comboBox2.Enabled = false;
                comboBox5.Enabled = false;
                checkBox5.Enabled = false;
                contextMenuStrip1.Enabled = false;


                cn.zagruzka_spiska_tp_energo();

            }



        }
        private void checkBox_v_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_v.Checked == true)
            { cn.zagruzka_spiska_tp_energo();
            cn.zagruzka_spiska_othod_tp_energo();

            checkBox_z.Enabled = false;
            checkBox_s.Enabled = false;
            checkBox_u.Enabled = false;
            }
            if (checkBox_v.Checked == false)
            {
                checkBox_z.Enabled = true;
                checkBox_s.Enabled = true;
                checkBox_u.Enabled = true;

            }

        }
        private void checkBox_z_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_z.Checked == true)
            {
                cn.zagruzka_spiska_tp_energo();
                cn.zagruzka_spiska_othod_tp_energo();

                checkBox_v.Enabled = false;
                checkBox_s.Enabled = false;
                checkBox_u.Enabled = false;
            }
            if (checkBox_z.Checked == false)
            {
                checkBox_v.Enabled = true;
                checkBox_s.Enabled = true;
                checkBox_u.Enabled = true;

            }

        }
        private void checkBox_s_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_s.Checked == true)
            {
                cn.zagruzka_spiska_tp_energo();
                cn.zagruzka_spiska_othod_tp_energo();

                checkBox_z.Enabled = false;
                checkBox_v.Enabled = false;
                checkBox_u.Enabled = false;
            }
            if (checkBox_s.Checked == false)
            {
                checkBox_z.Enabled = true;
                checkBox_v.Enabled = true;
                checkBox_u.Enabled = true;

            }
        }
        private void checkBox_u_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_u.Checked == true)
            {
                cn.zagruzka_spiska_tp_energo();
                cn.zagruzka_spiska_othod_tp_energo();

                checkBox_z.Enabled = false;
                checkBox_s.Enabled = false;
                checkBox_v.Enabled = false;
            }
            if (checkBox_u.Checked == false)
            {
                
                checkBox_z.Enabled = true;
                checkBox_s.Enabled = true;
                checkBox_v.Enabled = true;

            }
        }
        private void checkBox_all_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_all.Checked == true)
            {
                for (int i = 0; i < checkedListBox1.Items.Count; i++)
                {
                    checkedListBox1.SetItemChecked(i, true);
                }
            }

            if (checkBox_all.Checked == false)
            {
                for (int i = 0; i < checkedListBox1.Items.Count; i++)
                {
                    checkedListBox1.SetItemChecked(i, false);
                }
            }

        }
        private void textBox_poisk_TextChanged(object sender, EventArgs e)
        {
            cn.poisk_tp();
        }

    }
}
