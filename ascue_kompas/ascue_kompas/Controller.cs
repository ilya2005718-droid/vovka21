using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms.DataVisualization.Charting;
using System.Windows.Forms;
using System.Threading;
using ascue_kompas;
namespace ascue_kompas
{
    public class Controller
    {
        View v;
        Model m;
        public Controller(View v1, Model m1)
        {
            v = v1;
            m = m1;
            v.setController(this);
        }
        public string zagolovok;
        public void zagruzka_spiska_rp()
        {
            v.comboBox1 = m.spisok_rp(v.comboBox1);
        }
        public void logy()
        {
            m.log();
        }
        public void zagruzka_spiska_prisoed()
        {
            v.comboBox5 = m.spisok_othod_sutki(v.comboBox5,v.comboBox1);
        }
        public void zagruzka_spiska_othod_rp()
        {
            v.checkedListBox1 = m.spisok_othod_rp(v.checkedListBox1,v.comboBox1, v.cb_base);
        }
        public void zagruzka_spiska_othod_rp_energo()
        {
            v.checkedListBox1 = m.spisok_othod_rp_energo(v.checkedListBox1,v.comboBox1);
        }
        public void zagruzka_spiska_rp_energo()
        {
            v.comboBox1 = m.spisok_rp_energo(v.comboBox1);
        }
        public void zagruzka_resultat()
        {
            v.dataGridView3 = m.resultat(v.dataGridView3,v.comboBox1,v.checkedListBox1,v.dateTimePicker1,v.dateTimePicker2,v.checkBox4,v.comboBox3, v.comboBox4, v.contextMenuStrip1, v.richTextBox1, v.cb_base);
           
        }
        public void zagruzga_resultat_energomera()
        {
            v.dataGridView3 = m.resultat_energomera(v.dataGridView3, v.comboBox1, v.checkedListBox1,v.dateTimePicker1,v.dateTimePicker2,v.comboBox3,v.comboBox4, v.contextMenuStrip1, v.richTextBox1, v.checkBox4);
        }
        public void zagruzka_graphic()
        {
            v.chart1 = m.graphic(v.chart1, v.dataGridView3, 6, 5, 5, 4, "dd","кВт*ч","Дни", v.checkedListBox1, v.dateTimePicker1, v.dateTimePicker2, v.cb_base);
        }
        public void check_pok()
        {
            v.checkBox1 = m.videl_pok(v.checkBox1, v.checkedListBox1);
        }
        public void check_sutki()
        {
            v.checkBox2 = m.videl_tp(v.checkBox2, v.checkedListBox1);
        }
        public void check_p()
        {
            v.checkBox3 = m.videl_p(v.checkBox3, v.checkedListBox1);
        }
        public void export_to_excel()
        {
            SaveFileDialog svd1 = new SaveFileDialog();
            svd1.Filter = "xls files(*.xls)|*.xls|All files(*.*)|*.*";
            svd1.FilterIndex = 1;
            svd1.RestoreDirectory = true;
            if (svd1.ShowDialog() == DialogResult.OK)
            {
                m.file = svd1.FileName;
            }
            else return;
            m.export("рп", v.dataGridView3);
        }
        public void nebalans_v()
        {
           v.dataGridView3 =  m.nebalans_vis(v.dataGridView3, v.dateTimePicker1, v.dateTimePicker2, v.comboBox1,v.contextMenuStrip1, v.chart2, v.cb_base, v.richTextBox1);
        }
        public void nebalans_n()
        {
            v.dataGridView3 = m.nebalans_niz(v.dataGridView3, v.dateTimePicker1, v.dateTimePicker2, v.comboBox1, v.contextMenuStrip1, v.chart2, v.cb_base, v.richTextBox1);
        }
        public void sutochnoe_potreb()
        {
            v.dataGridView3 = m.sut_potreb(v.dataGridView3, v.dateTimePicker1, v.dateTimePicker2, v.comboBox1, v.contextMenuStrip1, v.chart2, v.cb_base, v.richTextBox1);
        }
        public void sutochniy_nebalans()
        {
            m.nebalans_sutki(v.dataGridView3, v.dataGridView1,v.checkBox5, v.comboBox1,v.dateTimePicker3, v.richTextBox1);
            m.nebalans_sutki1(v.dataGridView3, v.dataGridView1);
            m.nebalans_sutki_koef(v.dataGridView3, v.dataGridView1);
            m.nebalans_sutki_sum_columns(v.dataGridView3, v.dataGridView1);
            v.dataGridView3 = m.nebalans_sutki_finish(v.dataGridView3, v.dataGridView1, v.checkBox5, v.comboBox1, v.dateTimePicker3, v.chart3, v.contextMenuStrip1);
        }
        public void sutochniy_nebalans_energomera()
        {
            m.nebalans_sutki_energomera(v.dataGridView3, v.dataGridView1, v.checkBox5, v.comboBox1, v.dateTimePicker3, v.richTextBox1);
            m.nebalans_sutki1_energo(v.dataGridView3, v.dataGridView1);
            m.nebalans_sutki_koef(v.dataGridView3, v.dataGridView1);
            m.nebalans_sutki_sum_columns(v.dataGridView3, v.dataGridView1);
            v.dataGridView3 = m.nebalans_sutki_finish(v.dataGridView3, v.dataGridView1, v.checkBox5, v.comboBox1, v.dateTimePicker3, v.chart3, v.contextMenuStrip1);
        }
        public void sutochniy_nebalans_prisoed()
        {
            m.nebalans_sutki(v.dataGridView3, v.dataGridView1, v.checkBox5, v.comboBox1, v.dateTimePicker3, v.richTextBox1);
            m.nebalans_sutki1(v.dataGridView3, v.dataGridView1);
            m.nebalans_sutki_koef(v.dataGridView3, v.dataGridView1);
            v.dataGridView3 = m.nebalans_sutki_finish_prisoed(v.dataGridView3,v.dataGridView1);
        }
        public void sutochniy_nebalans_prisoed_energo()
        {
            m.nebalans_sutki_energomera(v.dataGridView3, v.dataGridView1, v.checkBox5, v.comboBox1, v.dateTimePicker3, v.richTextBox1);
            m.nebalans_sutki1_energo(v.dataGridView3, v.dataGridView1);
            m.nebalans_sutki_koef(v.dataGridView3, v.dataGridView1);
            v.dataGridView3 = m.nebalans_sutki_finish_prisoed(v.dataGridView3, v.dataGridView1);
        }
        public void sut_nebalans_povremeni()
        {
            v.dataGridView3 = m.nebalans_sutki_finish_prisoed_povremeni(v.dataGridView3, v.comboBox2);
        }
        public void sutochniy_nebalans_prisoed_poothod()
        {
            v.dataGridView3 = m.nebalans_sutki_finish_prisoed_poothod(v.dataGridView3,v.comboBox5, v.chart3, v.checkBox5, v.dateTimePicker3, v.contextMenuStrip1);

        }
        public void print_page(System.Drawing.Printing.PrintPageEventArgs e)
        {
            m.print_page(e,v.dataGridView3,v.comboBox1,zagolovok);
        }
        public void print_begin()
        {
            m.print_begin(v.dataGridView3);
        }
        public void vid_graphica()
        {
            m.vid_graphic(v.comboBox7, v.chart1, v.chart2, v.chart3);
        }
        public void vid_graphica_proc()
        {
            m.vid_graphic_proc_nebalans(v.chart2,v.comboBox6);
        }
        public void vid_graphica_proc_sutki()
        {
            m.vid_graphic_proc_sut(v.chart3,v.dataGridView3,v.comboBox8);
        }
        public void povernut_prisoed()
        {
            v.dataGridView3 = m.povernut_prisoed(v.dataGridView3);
        }
        public void nebalans_v_per()
        {
            v.dataGridView3 = m.nebalans_vis_per(v.dataGridView3, v.dateTimePicker1, v.dateTimePicker2, v.comboBox1, v.contextMenuStrip1, v.chart2, v.cb_base, v.richTextBox1);
        }
        public void report_neotvet_all()
        {
            v.dataGridView3 = m.report_neotvet(v.dataGridView3,v.comboBox1,v.checkedListBox1,v.dateTimePicker1,v.dateTimePicker2,v.comboBox3,v.comboBox4,v.richTextBox1, "Все");
        }
        public void report_neotvet_vis()
        {
            v.dataGridView3 = m.report_neotvet(v.dataGridView3, v.comboBox1, v.checkedListBox1, v.dateTimePicker1, v.dateTimePicker2, v.comboBox3, v.comboBox4, v.richTextBox1, "vis");
        }
        public void report_neotvet_niz()
        {
            v.dataGridView3 = m.report_neotvet(v.dataGridView3, v.comboBox1, v.checkedListBox1, v.dateTimePicker1, v.dateTimePicker2, v.comboBox3, v.comboBox4, v.richTextBox1, "niz");
        }
        public void report_neotvet_rezerv()
        {
            v.dataGridView3 = m.report_neotvet(v.dataGridView3, v.comboBox1, v.checkedListBox1, v.dateTimePicker1, v.dateTimePicker2, v.comboBox3, v.comboBox4, v.richTextBox1, "rezerv");
        }

        //Энергосфера

        public void zagruzka_spiska_r()
    {
        v.comboBox1 = m.spisok_r(v.comboBox1, v.richTextBox1);
    }
        public void zagruzka_spiska_othod_r()
        {
            v.checkedListBox1 = m.spisok_othod_r(v.checkedListBox1, v.comboBox1,v.richTextBox1);
        }
        public void zagruzka_resultat_sphera()
    {
        v.dataGridView3 = m.resultat_sphera(v.dataGridView3, v.comboBox1, v.checkedListBox1, v.dateTimePicker1, v.dateTimePicker2, v.comboBox3, v.comboBox4, v.richTextBox1);
    }

        //Меркурии
        
        public void zagruzka_spiska_tp_mercury()
        {
            v.comboBox1 = m.spisok_tp_mercury(v.comboBox1);
        }
        public void zagruzka_spiska_oth_tp_mercury()
        {
            v.checkedListBox1 = m.spisok_othod_tp_mercury(v.checkedListBox1,v.comboBox1,v.cb_base,v.richTextBox1);
        }
        public void zagruzka_resultat_mercury()
        {
            v.dataGridView3 = m.resultat_mercury(v.dataGridView3,v.comboBox1,v.checkedListBox1,v.dateTimePicker1,v.dateTimePicker2,v.richTextBox1);
        }

        //Энергомера ТП

        public void zagruzka_spiska_tp_energo()
        {
            v.comboBox1 = m.spisok_tp_energo(v.comboBox1, v.checkBox_v, v.checkBox_z, v.checkBox_s, v.checkBox_u, v.richTextBox1);
        }
        public void zagruzka_spiska_othod_tp_energo()
        {
            v.checkedListBox1 = m.spisok_othod_tp_energo(v.checkedListBox1, v.comboBox1, v.richTextBox1);
        }
        public void zagruzka_resultat_tp_energo()
        {
            v.dataGridView3 = m.resultat_tp_energo(v.dataGridView3,v.comboBox1, v.checkedListBox1, v.dateTimePicker1, v.dateTimePicker2, v.richTextBox1);
        }
        public void poisk_tp()
        {
            m.poisk_tp(v.textBox_poisk,v.comboBox1,v.checkedListBox1);
        }

    }
}
