using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Net.NetworkInformation;
using System.Xml;
using System.Drawing.Printing;
using Npgsql;
using ExcelLibrary;
using ExcelLibrary.SpreadSheet;
using ExcelLibrary.CompoundDocumentFormat;
using ExcelLibrary.BinaryFileFormat;
using ascue_kompas;
using System.Windows.Forms.DataVisualization.Charting;
using System.Collections;
using System.Threading;
using MySql.Data.MySqlClient;

namespace ascue_kompas
{
    public class Model
    {
        public string monitor = "Server=192.168.150.40;Port=5432;;User Id=psql; Password=;Database=monitor; CommandTimeout=955555";
        public string monitor_options = "Server=192.168.150.40;Port=5432;;User Id=psql; Password=;Database=monitor_options; CommandTimeout=955555";
        public string sql_connection = "Server = 192.168.144.221; Initial Catalog = ascue; Integrated Security=False; User ID = ascue; Password = ascue";
        public string energomera_comobjects = "Server=192.168.143.15;Port=5432;;User Id=bee; Password=123;Database=RGES; CommandTimeout=955555";
        public string put_sphera = "Server = 10.10.1.3; Initial Catalog = ASKUE; Integrated Security=False; User ID = creator; Password = 1; MultipleActiveResultSets=True";
        public string mercury = "Server=192.168.150.56;Port=5432;;User Id=postgres; Password=;Database=mercury; CommandTimeout=955555";
        public string energomera_tp_connection = "Server=192.168.143.15;Port=5432;;User Id=postgres; Password=postgres;Database=RGES; CommandTimeout=955555";

        public string mysql_connection = "Persist Security Info=False;database=ascue;server=192.168.144.221;port=3306;user id=ascue;Password=ascue321;encrypt=no;charset=utf8";

        public string file = String.Empty;
        public string mes1 = "";
        public string mes2 = "";
        public string ru_vis = " РУ - 6-10 кВ";
        public string ru_niz = " РУ - 0,4 кВ";
        public int prov = 1;
        public string ru = "";
        StringFormat strFormat; //Used to format the grid rows.
        ArrayList arrColumnLefts = new ArrayList();//Used to save left coordinates of columns
        ArrayList arrColumnWidths = new ArrayList();//Used to save column widths
        public int iCellHeight = 0; //Used to get/set the datagridview cell height
        public int iTotalWidth = 0; //
        public int iRow = 0;//Used as counter
        public bool bFirstPage = false; //Used to check whether we are printing first page
        public bool bNewPage = false;// Used to check whether we are printing a new page
        public int iHeaderHeight = 0; //Used for the header height

        class class0
        {
            public string t_time { get; set; }
            public string values_id { get; set; }
            public string v_value { get; set; }
            public string valid { get; set; }
        }
        class class0_energo
        {
            public string t_time { get; set; }
            public string meters_id { get; set; }
            public string v_value { get; set; }
            public string sn { get; set; }
        }
        class class1
        {
            public string t_time { get; set; }
            public string values_id { get; set; }
            public string v_value { get; set; }
            public string valid { get; set; }
        } //класс с названием столбцов для sql
        class class2
        {
            public string values_id { get; set; }
            public string name { get; set; }
            public string name_f { get; set; }
            public string tag { get; set; }
            // public string meas { get; set; }
        } //создаем класс с названием столбцов для mdb
        class class2_energo
        {
            public string meters_id { get; set; }
            public string name { get; set; }
            public string name_f { get; set; }
            public string sn { get; set; }
            // public string meas { get; set; }
        } //создаем класс с названием столбцов для mdb
        class class3
        {
            public string name { get; set; }
            public string name_f { get; set; }
            public string tag { get; set; }
            public string schet { get; set; }
            public string ktt { get; set; }
            public string t_time { get; set; }
            public string v_value { get; set; }
            public string valid { get; set; }
            //public string meas { get; set; }
        }
        class class3_1
        {
            public string name { get; set; }
            public string name_f { get; set; }
            public string tag { get; set; }
            public string t_time { get; set; }
            public string v_value { get; set; }
            public string valid { get; set; }
            //public string meas { get; set; }
        }
        class class4
        {
            public void v(int index, String p)
            {

                switch (index)
                {
                    case 0:
                        p1 = p;
                        break;
                    case 1:
                        p2 = p;
                        break;
                    case 2:
                        p3 = p;
                        break;
                    case 3:
                        p4 = p;
                        break;
                    case 4:
                        p5 = p;
                        break;
                    case 5:
                        p6 = p;
                        break;
                    case 6:
                        p7 = p;
                        break;
                    case 7:
                        p8 = p;
                        break;
                    case 8:
                        p9 = p;
                        break;
                    case 9:
                        p10 = p;
                        break;
                    case 10:
                        p11 = p;
                        break;
                    case 11:
                        p12 = p;
                        break;
                    case 12:
                        p13 = p;
                        break;
                    case 13:
                        p14 = p;
                        break;
                    case 14:
                        p15 = p;
                        break;
                    case 15:
                        p16 = p;
                        break;
                    case 16:
                        p17 = p;
                        break;
                    case 17:
                        p18 = p;
                        break;
                    case 18:
                        p19 = p;
                        break;
                    case 19:
                        p20 = p;
                        break;
                    case 20:
                        p21 = p;
                        break;
                    case 21:
                        p22 = p;
                        break;
                    case 22:
                        p23 = p;
                        break;
                    case 23:
                        p24 = p;
                        break;
                    case 24:
                        p25 = p;
                        break;
                    case 25:
                        p26 = p;
                        break;
                    case 26:
                        p27 = p;
                        break;
                    case 27:
                        p28 = p;
                        break;
                    case 28:
                        p29 = p;
                        break;
                    case 29:
                        p30 = p;
                        break;
                    case 30:
                        p31 = p;
                        break;
                    case 31:
                        p32 = p;
                        break;
                    case 32:
                        p33 = p;
                        break;
                    case 33:
                        p34 = p;
                        break;
                    case 34:
                        p35 = p;
                        break;
                    case 35:
                        p36 = p;
                        break;
                    case 36:
                        p37 = p;
                        break;
                    case 37:
                        p38 = p;
                        break;
                    case 38:
                        p39 = p;
                        break;
                    case 39:
                        p40 = p;
                        break;
                    case 40:
                        p41 = p;
                        break;
                    case 41:
                        p42 = p;
                        break;
                    case 42:
                        p43 = p;
                        break;
                    case 43:
                        p44 = p;
                        break;
                    case 44:
                        p45 = p;
                        break;
                    case 45:
                        p46 = p;
                        break;
                    case 46:
                        p47 = p;
                        break;
                    case 47:
                        p48 = p;
                        break;
                    case 48:
                        p49 = p;
                        break;

                }
            }
            public string name_rp { get; set; }
            public string p1 { get; set; }
            public string p2 { get; set; }
            public string p3 { get; set; }
            public string p4 { get; set; }
            public string p5 { get; set; }
            public string p6 { get; set; }
            public string p7 { get; set; }
            public string p8 { get; set; }
            public string p9 { get; set; }
            public string p10 { get; set; }
            public string p11 { get; set; }
            public string p12 { get; set; }
            public string p13 { get; set; }
            public string p14 { get; set; }
            public string p15 { get; set; }
            public string p16 { get; set; }
            public string p17 { get; set; }
            public string p18 { get; set; }
            public string p19 { get; set; }
            public string p20 { get; set; }
            public string p21 { get; set; }
            public string p22 { get; set; }
            public string p23 { get; set; }
            public string p24 { get; set; }
            public string p25 { get; set; }
            public string p26 { get; set; }
            public string p27 { get; set; }
            public string p28 { get; set; }
            public string p29 { get; set; }
            public string p30 { get; set; }
            public string p31 { get; set; }
            public string p32 { get; set; }
            public string p33 { get; set; }
            public string p34 { get; set; }
            public string p35 { get; set; }
            public string p36 { get; set; }
            public string p37 { get; set; }
            public string p38 { get; set; }
            public string p39 { get; set; }
            public string p40 { get; set; }
            public string p41 { get; set; }
            public string p42 { get; set; }
            public string p43 { get; set; }
            public string p44 { get; set; }
            public string p45 { get; set; }
            public string p46 { get; set; }
            public string p47 { get; set; }
            public string p48 { get; set; }
            public string p49 { get; set; }

        }
        class class4_1
        {
            public void v(int index, String p)
            {

                switch (index)
                {
                    case 0:
                        p1 = p;
                        break;
                    case 1:
                        p2 = p;
                        break;
                    case 2:
                        p3 = p;
                        break;
                    case 3:
                        p4 = p;
                        break;
                    case 4:
                        p5 = p;
                        break;
                    case 5:
                        p6 = p;
                        break;
                    case 6:
                        p7 = p;
                        break;
                    case 7:
                        p8 = p;
                        break;
                    case 8:
                        p9 = p;
                        break;
                    case 9:
                        p10 = p;
                        break;
                    case 10:
                        p11 = p;
                        break;
                    case 11:
                        p12 = p;
                        break;
                    case 12:
                        p13 = p;
                        break;
                    case 13:
                        p14 = p;
                        break;
                    case 14:
                        p15 = p;
                        break;
                    case 15:
                        p16 = p;
                        break;
                    case 16:
                        p17 = p;
                        break;
                    case 17:
                        p18 = p;
                        break;
                    case 18:
                        p19 = p;
                        break;
                    case 19:
                        p20 = p;
                        break;
                    case 20:
                        p21 = p;
                        break;
                    case 21:
                        p22 = p;
                        break;
                    case 22:
                        p23 = p;
                        break;
                    case 23:
                        p24 = p;
                        break;
                    case 24:
                        p25 = p;
                        break;
                    case 25:
                        p26 = p;
                        break;
                    case 26:
                        p27 = p;
                        break;
                    case 27:
                        p28 = p;
                        break;
                    case 28:
                        p29 = p;
                        break;
                    case 29:
                        p30 = p;
                        break;
                    case 30:
                        p31 = p;
                        break;
                    case 31:
                        p32 = p;
                        break;
                    case 32:
                        p33 = p;
                        break;
                    case 33:
                        p34 = p;
                        break;
                    case 34:
                        p35 = p;
                        break;
                    case 35:
                        p36 = p;
                        break;
                    case 36:
                        p37 = p;
                        break;
                    case 37:
                        p38 = p;
                        break;
                    case 38:
                        p39 = p;
                        break;
                    case 39:
                        p40 = p;
                        break;
                    case 40:
                        p41 = p;
                        break;
                    case 41:
                        p42 = p;
                        break;
                    case 42:
                        p43 = p;
                        break;
                    case 43:
                        p44 = p;
                        break;
                    case 44:
                        p45 = p;
                        break;
                    case 45:
                        p46 = p;
                        break;
                    case 46:
                        p47 = p;
                        break;
                    case 47:
                        p48 = p;
                        break;
                    case 48:
                        p49 = p;
                        break;

                }
            }
            public string name_rp { get; set; }
            public string p1 { get; set; }
            public string p2 { get; set; }
            public string p3 { get; set; }
            public string p4 { get; set; }
            public string p5 { get; set; }
            public string p6 { get; set; }
            public string p7 { get; set; }
            public string p8 { get; set; }
            public string p9 { get; set; }
            public string p10 { get; set; }
            public string p11 { get; set; }
            public string p12 { get; set; }
            public string p13 { get; set; }
            public string p14 { get; set; }
            public string p15 { get; set; }
            public string p16 { get; set; }
            public string p17 { get; set; }
            public string p18 { get; set; }
            public string p19 { get; set; }
            public string p20 { get; set; }
            public string p21 { get; set; }
            public string p22 { get; set; }
            public string p23 { get; set; }
            public string p24 { get; set; }
            public string p25 { get; set; }
            public string p26 { get; set; }
            public string p27 { get; set; }
            public string p28 { get; set; }
            public string p29 { get; set; }
            public string p30 { get; set; }
            public string p31 { get; set; }
            public string p32 { get; set; }
            public string p33 { get; set; }
            public string p34 { get; set; }
            public string p35 { get; set; }
            public string p36 { get; set; }
            public string p37 { get; set; }
            public string p38 { get; set; }
            public string p39 { get; set; }
            public string p40 { get; set; }
            public string p41 { get; set; }
            public string p42 { get; set; }
            public string p43 { get; set; }
            public string p44 { get; set; }
            public string p45 { get; set; }
            public string p46 { get; set; }
            public string p47 { get; set; }
            public string p48 { get; set; }
            public string p49 { get; set; }

        }
        class class5
        {

            public string fider { get; set; }
            public string val { get; set; }
        }
        class class5_energo
        {

            public string fider { get; set; }
            public string val { get; set; }
        }
        class class6
        {

            public string name_tii { get; set; }
            public string num_schet { get; set; }
        }  //класс с названием фидера и счетчика то есть для ввода номеров сч
        class class7
        {

            public string name_tii { get; set; }
            public string ktt { get; set; }
        }
        class class8
        {
            public string name_f { get; set; }
            public string koef { get; set; }
        }
        class class8_1
        {
            public string name_f { get; set; }
            public string koef { get; set; }
        }
        class class9
        {
            public string t_time { get; set; }
            public string values_id { get; set; }
            public string v_value { get; set; }
            public string valid { get; set; }
        }
        class class9_energo
        {
            public string t_time { get; set; }
            public string meters_id { get; set; }
            public string v_value { get; set; }
            public string sn { get; set; }
        }
        class class10
        {

            public string rp { get; set; }
            //public string vvod { get; set; }
            public string data { get; set; }
            public string val_f { get; set; }
            public string val_v { get; set; }
            public string raznica { get; set; }
            public string proc { get; set; }


        }
        class class_sutki
        {
            public string id { get; set; }
            public string name { get; set; }
            public decimal value { get; set; }
            public long data { get; set; }
        }
        class class_sutki_energo
        {
            public string id { get; set; }
            public string name { get; set; }
            public decimal value { get; set; }
            public byte data { get; set; }
        }
        class class_finish
        {
            public string name { get; set; }
            public string name_f { get; set; }
            public string tag { get; set; }
            public string schet { get; set; }
            public string ktt { get; set; }
            public string t_time { get; set; }
            public string v_value { get; set; }
            public string valid { get; set; }
        }
        class class_finish_energo
        {
            public string name { get; set; }
            public string name_f { get; set; }
            public string schet { get; set; }
            public string ktt { get; set; }
            public string t_time { get; set; }
            public string v_value { get; set; }
        }
        class energomera
        {
            public string name_rp { get; set; }
            public string name_f { get; set; }
            public string name_sch { get; set; }
            public string ktt { get; set; }
            public string t_time { get; set; }
            public string v_value { get; set; }
        }
        class energomera_v
        {
            public string name_rp { get; set; }
            public string name_f { get; set; }
            public string name_sch { get; set; }
            //public string ktt { get; set; }
            public string t_time { get; set; }
            public string v_value { get; set; }
        }
        class class_vrem_tabl
        {
            public int id { get; set; }
            public int type { get; set; }
            public int num { get; set; }
        }
        class class_nebalans
        {
            public string data { get; set; }
            public string value { get; set; }
            public string proc { get; set; }
        }
        class class_sut_potreb
        {
            public string data { get; set; }
            public string sum { get; set; }
            public string value { get; set; }
            public string proc { get; set; }
        }
        class chas
        {
            public string name_f { get; set; }
            public string t_time { get; set; }
            public string v_value { get; set; }
        }
        class one_fider
        {
            public string name_f { get; set; }
            public string t_time { get; set; }
            public string v_value { get; set; }
        }
        class povernut
        {
            public void v(int index, String p)
            {

                switch (index)
                {
                    case 0:
                        p1 = p;
                        break;
                    case 1:
                        p2 = p;
                        break;
                    case 2:
                        p3 = p;
                        break;
                    case 3:
                        p4 = p;
                        break;
                    case 4:
                        p5 = p;
                        break;
                    case 5:
                        p6 = p;
                        break;
                    case 6:
                        p7 = p;
                        break;
                    case 7:
                        p8 = p;
                        break;
                    case 8:
                        p9 = p;
                        break;
                    case 9:
                        p10 = p;
                        break;
                    case 10:
                        p11 = p;
                        break;
                    case 11:
                        p12 = p;
                        break;
                    case 12:
                        p13 = p;
                        break;
                    case 13:
                        p14 = p;
                        break;
                    case 14:
                        p15 = p;
                        break;
                    case 15:
                        p16 = p;
                        break;
                    case 16:
                        p17 = p;
                        break;
                    case 17:
                        p18 = p;
                        break;
                    case 18:
                        p19 = p;
                        break;
                    case 19:
                        p20 = p;
                        break;
                    case 20:
                        p21 = p;
                        break;
                    case 21:
                        p22 = p;
                        break;
                    case 22:
                        p23 = p;
                        break;
                    case 23:
                        p24 = p;
                        break;
                    case 24:
                        p25 = p;
                        break;
                    case 25:
                        p26 = p;
                        break;
                    case 26:
                        p27 = p;
                        break;
                    case 27:
                        p28 = p;
                        break;
                    case 28:
                        p29 = p;
                        break;
                    case 29:
                        p30 = p;
                        break;
                    case 30:
                        p31 = p;
                        break;
                    case 31:
                        p32 = p;
                        break;
                    case 32:
                        p33 = p;
                        break;
                    case 33:
                        p34 = p;
                        break;
                    case 34:
                        p35 = p;
                        break;
                    case 35:
                        p36 = p;
                        break;
                    case 36:
                        p37 = p;
                        break;
                    case 37:
                        p38 = p;
                        break;
                    case 38:
                        p39 = p;
                        break;
                    case 39:
                        p40 = p;
                        break;
                    case 40:
                        p41 = p;
                        break;
                    case 41:
                        p42 = p;
                        break;
                    case 42:
                        p43 = p;
                        break;
                    case 43:
                        p44 = p;
                        break;
                    case 44:
                        p45 = p;
                        break;
                    case 45:
                        p46 = p;
                        break;
                    case 46:
                        p47 = p;
                        break;
                    case 47:
                        p48 = p;
                        break;

                }
            }
            public string name_rp { get; set; }
            public string p1 { get; set; }
            public string p2 { get; set; }
            public string p3 { get; set; }
            public string p4 { get; set; }
            public string p5 { get; set; }
            public string p6 { get; set; }
            public string p7 { get; set; }
            public string p8 { get; set; }
            public string p9 { get; set; }
            public string p10 { get; set; }
            public string p11 { get; set; }
            public string p12 { get; set; }
            public string p13 { get; set; }
            public string p14 { get; set; }
            public string p15 { get; set; }
            public string p16 { get; set; }
            public string p17 { get; set; }
            public string p18 { get; set; }
            public string p19 { get; set; }
            public string p20 { get; set; }
            public string p21 { get; set; }
            public string p22 { get; set; }
            public string p23 { get; set; }
            public string p24 { get; set; }
            public string p25 { get; set; }
            public string p26 { get; set; }
            public string p27 { get; set; }
            public string p28 { get; set; }
            public string p29 { get; set; }
            public string p30 { get; set; }
            public string p31 { get; set; }
            public string p32 { get; set; }
            public string p33 { get; set; }
            public string p34 { get; set; }
            public string p35 { get; set; }
            public string p36 { get; set; }
            public string p37 { get; set; }
            public string p38 { get; set; }
            public string p39 { get; set; }
            public string p40 { get; set; }
            public string p41 { get; set; }
            public string p42 { get; set; }
            public string p43 { get; set; }
            public string p44 { get; set; }
            public string p45 { get; set; }
            public string p46 { get; set; }
            public string p47 { get; set; }
            public string p48 { get; set; }


        }
        class res_sphera
        {
            public string name_r { get; set; }
            public string name_l { get; set; }
            public string koef { get; set; }
            public string time { get; set; }
            public string val_pok { get; set; }
        }
        class res_mercury
        {
            public string name_tp { get; set; }
            public string name_t { get; set; }
            public string s_number { get; set; }
            public string data { get; set; }
            public string tek_val { get; set; }
        }
        class res_tp_energo
        {
            public string name_tp { get; set; }
            public string s_number { get; set; }
            public string data { get; set; }
            public string val { get; set; }
            
        }

        List<class2> list1 = new List<class2>();  //создаем лист для sql
        List<class2_energo> list1_energo = new List<class2_energo>();
        List<class5> fider = new List<class5>(); //для загрузки ктт
        List<class5_energo> energo_ktt = new List<class5_energo>(); //для загрузки ктт энергомера
        List<class6> schet = new List<class6>(); // лист для загрузки номеров счетчиков
        List<class_sutki> list_sutki = new List<class_sutki>();
        List<class_sutki_energo> list_sutki_energo = new List<class_sutki_energo>();
        List<class3> result = new List<class3>(); //создаем лист для конечного результата
        List<class3_1> result_v = new List<class3_1>();
        List<class0> list = new List<class0>();    //создаем лист для mdb
        List<class0_energo> list_energo = new List<class0_energo>();
        List<class7> listk = new List<class7>();
        List<class_finish> finish = new List<class_finish>(); //лист для вывода конечных данных
        List<class_finish_energo> finish_energo = new List<class_finish_energo>();
        List<class_vrem_tabl> vrem_tabl = new List<class_vrem_tabl>();
        List<string> list_v = new List<string>(); //лист вводов высокой
        List<string> list_f = new List<string>(); //лист фидеров высокой
        List<string> list_vn = new List<string>();
        List<string> list_fn = new List<string>();
        List<string> list_v_niz = new List<string>(); //лист вводов низкой
        List<string> list_f_niz = new List<string>(); //лист фидеров низкой
        List<string> list_vn_niz = new List<string>();
        List<string> list_fn_niz = new List<string>();
        List<decimal> b = new List<decimal>();
        List<decimal> b1 = new List<decimal>();
        List<DateTime> dat = new List<DateTime>();
        List<class_nebalans> list_nebalans = new List<class_nebalans>();
        List<class_sut_potreb> list_sut_potreb = new List<class_sut_potreb>();
        List<class8> fider_koef = new List<class8>();
        List<class8_1> vvod_koef = new List<class8_1>();
        List<class9> list_vlist = new List<class9>();
        List<class9_energo> list_vlist_energo = new List<class9_energo>();
        List<class4> asc = new List<class4>();  //создаем лист для горизонтального вида
        List<class4_1> asc_v = new List<class4_1>();
        List<class10> trans = new List<class10>();
        List<chas> gr = new List<chas>();
        List<one_fider> of = new List<one_fider>();
        List<povernut> asc_p = new List<povernut>();
        List<energomera> vivod_energomera = new List<energomera>();
        List<energomera_v> vivod_energomera_v = new List<energomera_v>();
        List<res_sphera> vivod_sphera = new List<res_sphera>();
        List<res_mercury> vivod_mercury = new List<res_mercury>();
        List<res_tp_energo> vivod_tp_energo = new List<res_tp_energo>();
 

        public string tii(CheckedListBox clb)
        {
            string s = "";
            if (clb.CheckedItems.Count == 0) { MessageBox.Show("Не выбрано ни одного присоединения! \nВыберите присоединения и нажмите \"Результат\""); }
            for (int i = 0; i < clb.CheckedItems.Count; i++)
            {
                if (i == clb.CheckedItems.Count - 1) { s = s + "'" + clb.CheckedItems[i].ToString() + "'"; }
                else
                {
                    s = s + "'" + clb.CheckedItems[i].ToString() + "'" + ",";
                }
            }
            return s;
        } // для формирования запроса выбора выделенных элементов из чекбокслиста
        public void export(String namesheet, DataGridView nametabl)
        {
            Workbook workbook = new Workbook();
            Worksheet worksheet = new Worksheet(namesheet); //создаём новый лист

            for (int j = 1; j <= nametabl.Rows.Count; j++)
            {
                for (int i = 1; i <= nametabl.Columns.Count; i++)
                {
                    worksheet.Cells[j, i] = new Cell(nametabl.Rows[j - 1].Cells[i - 1].Value);
                }
            }

            workbook.Worksheets.Add(worksheet);
            workbook.Save(file);
        }
        public void ktt(ComboBox cb_rp)
        {
            listk.Clear();

            NpgsqlConnection connection = new NpgsqlConnection(monitor_options);
            connection.Open();
            NpgsqlCommand com = new NpgsqlCommand("select \"name_tii\", \"ktt\" from \"balans\" where \"name_rp\" = '" + cb_rp.SelectedItem.ToString() + "'", connection);
            NpgsqlDataReader datar = com.ExecuteReader();

            while (datar.Read())
            {
                class7 m = new class7();
                m.name_tii = datar.GetValue(0).ToString();
                m.ktt = datar.GetValue(1).ToString();
                listk.Add(m);

            }
            connection.Close();
        }
        public decimal summ(String id, long iter, string vn, DateTimePicker dtp_n, ComboBox cb_base, RichTextBox rtb) //для суточного потребления
        {
            decimal koef = 0;
            // byte den = 1;
            decimal s = 0;
            for (int k = 0; k < listk.Count; k++)
            {
                if (listk[k].name_tii.Equals(vn)) { koef = Convert.ToDecimal(listk[k].ktt); }
            }


            DateTime d1 = new DateTime(dtp_n.Value.Year, dtp_n.Value.Month, dtp_n.Value.Day); //дата начала периода
            double x1 = d1.ToOADate();    //перевод нормальной даты в абсолютную

            long base_tick = new DateTime(1601, 1, 1).Ticks;
            long n_tick = d1.Ticks;
            long n_tick_mysql = n_tick - base_tick;
            
            long oneday = 864000000000;

            if (cb_base.SelectedItem.ToString() == "Компас РП") //для компаса
            {

                //SqlConnection con = new SqlConnection(sql_connection);
                MySqlConnection con = new MySqlConnection(mysql_connection);
                con.Open();
                //SqlCommand com5 = new SqlCommand("SELECT max(V_VALUE) - min(V_VALUE) FROM ti  where VALUES_ID = " + id + " and T_TIME >= " + (x1 + iter) + " and T_TIME <= " + (x1 + iter + 1), con);
                MySqlCommand com5 = new MySqlCommand("SELECT max(V_VALUE) - min(V_VALUE) FROM ti  where VALUES_ID = " + id + " and T_TIME >= " + (n_tick_mysql + iter) + " and T_TIME <= " + (n_tick_mysql + iter + oneday), con);
                // MessageBox.Show("SELECT max(V_VALUE) - min(V_VALUE) FROM ti  where VALUES_ID = " + id + " and T_TIME >= " + (x1 + iter) + " and T_TIME <= " + (x1 + iter + 1));
                //rtb.Text = rtb.Text + "SELECT max(V_VALUE) - min(V_VALUE) FROM ti  where VALUES_ID = " + id + " and T_TIME >= " + (x1 + iter) + " and T_TIME <= " + (x1 + iter + 1) + " \n \n";
                rtb.Text = rtb.Text + "SELECT max(V_VALUE) - min(V_VALUE) FROM ti  where VALUES_ID = " + id + " and T_TIME >= " + (n_tick_mysql + iter) + " and T_TIME <= " + (n_tick_mysql + iter + oneday) + " \n \n";
                com5.ExecuteNonQuery();
                s = Convert.ToDecimal(com5.ExecuteScalar());
                //MessageBox.Show(s.ToString());
                con.Close();
            }

            if (cb_base.SelectedItem.ToString() == "Энергомера") //для энергомеры
            {
                NpgsqlConnection connection = new NpgsqlConnection(energomera_comobjects);
                connection.Open();
                NpgsqlCommand com = new NpgsqlCommand("SELECT sum(\"ValueProfiles\".\"Val\") from  \"ValueProfiles\", \"Meters\" where \"ValueProfiles\".\"MeterId\" = " + id + " and \"ValueProfiles\".\"MeterId\"=\"Meters\".\"MeterId\" and \"ValueProfiles\".\"StateId\" = 1 and \"ValueProfiles\".\"DT\" >= '" + DateTime.FromOADate(x1 + iter) + "' and \"ValueProfiles\".\"DT\" < '" + DateTime.FromOADate(x1 + iter + 1) + "'", connection);
                //rtb.Text = rtb.Text + "SELECT max(\"Values\".\"Val\") - min(\"Values\".\"Val\") from  \"Values\", \"Meters\" where \"Values\".\"MeterId\" = " + id + " and \"Values\".\"MeterId\"=\"Meters\".\"MeterId\" and \"Values\".\"TariffId\" = 1 and \"Values\".\"DT\" >= '" + DateTime.FromOADate(x1 + iter) + "' and \"Values\".\"DT\" <= '" + DateTime.FromOADate(x1 + iter + 1) + "' \n \n";
                rtb.Text = rtb.Text + "SELECT sum(\"ValueProfiles\".\"Val\") from  \"ValueProfiles\", \"Meters\" where \"ValueProfiles\".\"MeterId\" = " + id + " and \"ValueProfiles\".\"MeterId\"=\"Meters\".\"MeterId\" and \"ValueProfiles\".\"StateId\" = 1 and \"ValueProfiles\".\"DT\" >= '" + DateTime.FromOADate(x1 + iter) + "' and \"ValueProfiles\".\"DT\" < '" + DateTime.FromOADate(x1 + iter + 1) + "' \n \n";
                //MessageBox.Show(rtb.Text);
                com.ExecuteNonQuery();
                try
                {
                    s = Convert.ToDecimal(com.ExecuteScalar()) / 2;
                }
                catch { }
                //s = Convert.ToDecimal(com.ExecuteScalar());
                //MessageBox.Show(s.ToString() + "---" + id);
                connection.Close();
            }


            return s * koef;
        }
        public decimal summ_period(String id, string vn, DateTimePicker dtp_n, DateTimePicker dtp_k, ComboBox cb_base, RichTextBox rtb) //для расчета небалансов
        {
            decimal koef = 0;
            decimal s = 0;
            // byte den = 1;
            for (int k = 0; k < listk.Count; k++)
            {
                if (listk[k].name_tii.Equals(vn)) { koef = Convert.ToDecimal(listk[k].ktt); }
            }


            DateTime d1 = new DateTime(dtp_n.Value.Year, dtp_n.Value.Month, dtp_n.Value.Day); //дата начала периода
            DateTime d2 = new DateTime(dtp_k.Value.Year, dtp_k.Value.Month, dtp_k.Value.Day);
            double x1 = d1.ToOADate();    //перевод нормальной даты в абсолютную
            double x2 = d2.ToOADate();

            long base_tick = new DateTime(1601, 1, 1).Ticks;
            long n_tick = d1.Ticks;
            long n_tick_mysql = n_tick - base_tick;
            long k_tick = d2.Ticks;
            long k_tick_mysql = k_tick - base_tick;
            long oneday = 864000000000;

            if (cb_base.SelectedItem.ToString() == "Компас РП") //для компаса
            {

                //SqlConnection con = new SqlConnection(sql_connection);
                MySqlConnection con = new MySqlConnection(mysql_connection);
                con.Open();
                //SqlCommand com5 = new SqlCommand("SELECT max(V_VALUE) - min(V_VALUE) FROM ti  where VALUES_ID = " + id + " and T_TIME >= " + x1 + " and T_TIME <= " + (x2+1), con);
                //SqlCommand com5 = new SqlCommand("SELECT t2.V_VALUE - t1.V_VALUE FROM ti t1, ti t2  where t1.VALUES_ID = " + id + " and t1.T_TIME = " + x1 + " and t2.T_TIME = " + (x2 + 1) + " and t1.VALUES_ID=t2.VALUES_ID", con);
                MySqlCommand com5 = new MySqlCommand("SELECT t2.V_VALUE - t1.V_VALUE FROM ti t1, ti t2  where t1.VALUES_ID = " + id + " and t1.T_TIME = " + n_tick_mysql + " and t2.T_TIME = " + (k_tick_mysql + oneday) + " and t1.VALUES_ID=t2.VALUES_ID", con);
                rtb.Text = rtb.Text + "SELECT t2.V_VALUE - t1.V_VALUE FROM ti t1, ti t2  where t1.VALUES_ID = " + id + " and t1.T_TIME = " + n_tick_mysql + " and t2.T_TIME = " + (k_tick_mysql + oneday) + " and t1.VALUES_ID=t2.VALUES_ID \n \n";
                com5.ExecuteNonQuery();
                s = Convert.ToDecimal(com5.ExecuteScalar());
                con.Close();
            }

            if (cb_base.SelectedItem.ToString() == "Энергомера") //для энергомеры
            {
                NpgsqlConnection connection = new NpgsqlConnection(energomera_comobjects);
                connection.Open();
                NpgsqlCommand com = new NpgsqlCommand("SELECT sum(\"ValueProfiles\".\"Val\") from  \"ValueProfiles\", \"Meters\" where \"ValueProfiles\".\"MeterId\" = " + id + " and \"ValueProfiles\".\"MeterId\"=\"Meters\".\"MeterId\" and \"ValueProfiles\".\"StateId\" = 1 and \"ValueProfiles\".\"DT\" >= '" + DateTime.FromOADate(x1) + "' and \"ValueProfiles\".\"DT\" < '" + DateTime.FromOADate(x2 + 1) + "'", connection);
                // rtb.Text = rtb.Text + "SELECT max(t2.\"Val\") - min(t1.\"Val\") from  \"Values\" t1, \"Values\" t2, \"Meters\" where t1.\"MeterId\" = " + id + " and t1.\"MeterId\"=\"Meters\".\"MeterId\" and t1.\"TariffId\" = 1 and t1.\"DT\" = '" + DateTime.FromOADate(x1) + "' and t2.\"DT\" = '" + DateTime.FromOADate(x2 + 1) + "' and t1.\"MeterId\" = t2.\"MeterId\" \n \n";
                rtb.Text = rtb.Text + "SELECT sum(\"ValueProfiles\".\"Val\") from  \"ValueProfiles\", \"Meters\" where \"ValueProfiles\".\"MeterId\" = " + id + " and \"ValueProfiles\".\"MeterId\"=\"Meters\".\"MeterId\" and \"ValueProfiles\".\"StateId\" = 1 and \"ValueProfiles\".\"DT\" >= '" + DateTime.FromOADate(x1) + "' and \"ValueProfiles\".\"DT\" < '" + DateTime.FromOADate(x2 + 1) + "' \n \n";
                com.ExecuteNonQuery();
                try
                {
                    s = Convert.ToDecimal(com.ExecuteScalar()) / 2;
                }
                catch { s = 0; } //иногда есть значения с E поэтому заменяем на ноль
                // MessageBox.Show(s.ToString() + "----" + id.ToString());
                connection.Close();
            }

            return s * koef; //умножаем на коэффициент
        }
        public int timeparser(String s) //процедура просмотра времени для определения получаса, что если минуты начинаются с цыфры 3, то это получас
        {
            int resultset = 0;
            int hour = int.Parse(s.Substring(0, s.IndexOf(":")));
            string minute = s.Substring(s.IndexOf(":") + 1, 1);
            resultset = hour * 2;
            if (minute.Equals("3")) { resultset++; }
            if (resultset == 0)
            {
                if (prov % 2 != 0) { } else { resultset = 48; } prov++;
            }
            return resultset;
        }
        public int timeparser_pris(String s) //процедура просмотра времени для определения получаса, что если минуты начинаются с цыфры 3, то это получас
        {
            int resultset = 0;
            int hour = int.Parse(s.Substring(0, s.IndexOf(":")));
            string minute = s.Substring(s.IndexOf(":") + 1, 1);
            resultset = hour * 2;
            if (minute.Equals("3")) { resultset++; }
            return resultset;
        }
        public void log()
        {
            DateTime data = DateTime.Now;  //Время запуска программы
            String host = System.Net.Dns.GetHostName(); //
            System.Net.IPAddress ip = System.Net.Dns.GetHostByName(host).AddressList[0];
            String conn_param = "Server=192.168.150.40;Port=5432;;User Id=psql; Password=;Database=Base; CommandTimeout=955555"; ;
            NpgsqlConnection conn = new NpgsqlConnection(conn_param);
            conn.Open();
            NpgsqlCommand commm = new NpgsqlCommand("insert into log_ascue_viewer values ('" + ip.ToString() + "', '" + host.ToString() + "', '" + data.ToString() + "')", conn);
            commm.ExecuteNonQuery();
            conn.Close();

        }
        public ComboBox spisok_rp(ComboBox cb)
        {
            //dtp_now.Value = DateTime.Now;
            NpgsqlConnection connection = new NpgsqlConnection(monitor);
            cb.Items.Clear();
            cb.Items.Clear();
            connection.Open();
            //NpgsqlCommand com = new NpgsqlCommand("select \"Name\" from \"GROUPS\" where \"Parent\" is null order by substring(\"Name\",3,3)::numeric ", connection);
            NpgsqlCommand com = new NpgsqlCommand("select \"Name\" from \"GROUPS\" where \"Parent\" is null order by regexp_replace(\"Name\", '[^0-9]', '', 'g')::numeric ", connection);
            NpgsqlDataReader datar = com.ExecuteReader();
            while (datar.Read()) { cb.Items.Add(datar.GetString(0)); }
            cb.SelectedItem = "РП1";
            connection.Close();

            return cb;

        } //Загрузка списка РП
        public ComboBox spisok_rp_energo(ComboBox cb)
        {
            //dtp_now.Value = DateTime.Now;
            NpgsqlConnection connection = new NpgsqlConnection(energomera_comobjects);
            cb.Items.Clear();
            cb.Items.Clear();
            connection.Open();
            NpgsqlCommand com = new NpgsqlCommand("select \"Name\" from \"ComObjects\" where \"Name\" like '%РП%' ", connection);
            NpgsqlDataReader datar = com.ExecuteReader();
            while (datar.Read()) { cb.Items.Add(datar.GetString(0)); }
            cb.SelectedItem = "Ю РП-76";
            connection.Close();

            return cb;
        } //список рп для энергомеры
        public ComboBox spisok_othod_sutki(ComboBox cb_prisoed, ComboBox cb_rp)
        {
            NpgsqlConnection connection1 = new NpgsqlConnection(monitor_options);
            connection1.Open();
            NpgsqlCommand com1 = new NpgsqlCommand("select \"name_tii\" from balans where \"name_rp\" = '" + cb_rp.SelectedItem + "' and (\"fider\" = TRUE or \"fider_niz\" = TRUE or \"vvod\" = TRUE or \"vvod_niz\" = TRUE ) ", connection1);
            NpgsqlDataReader datar1 = com1.ExecuteReader();
            cb_prisoed.Items.Clear();
            while (datar1.Read())
            {
                cb_prisoed.Items.Add(datar1.GetString(0));
            }
            connection1.Close();
            return cb_prisoed;
        }
        public CheckedListBox spisok_othod_rp(CheckedListBox clb, ComboBox cb, ComboBox cb_base)
        {
            string z = "";
            List<string> listok = new List<string>();
            NpgsqlConnection connection = new NpgsqlConnection(monitor);
            //NpgsqlConnection connection1 = new NpgsqlConnection(monitor_options);
            clb.Items.Clear();
            connection.Open();
            //connection1.Open();
            if (cb_base.SelectedItem.ToString() == "Компас РП") { z = "select t4.\"Name\" from (select t1.\"Groups_ID\", t2.\"Name\" from \"GROUPS\" t1, \"GROUPS\" t2 where t1.\"Tag\" = 'ТИИ' and t1.\"Parent\"=t2.\"Groups_ID\" and t2.\"Name\" = '" + cb.SelectedItem.ToString() + "') as t0, \"TB_VALUES\" t4 where t4.\"Groups_ID\"=t0.\"Groups_ID\" and t4.\"Tag\" like 'ТИ%'"; }
            if (cb_base.SelectedItem.ToString() == "Компас РП измер") { z = "select t4.\"Name\" from (select t1.\"Groups_ID\" from \"GROUPS\" t1 where t1.\"Name\" = '" + cb.SelectedItem.ToString() + "') as t0, \"TB_VALUES\" t4 where t4.\"Groups_ID\"=t0.\"Groups_ID\" and t4.\"Tag\" like 'ТИ%'"; }
            NpgsqlCommand command = new NpgsqlCommand(z, connection);
            NpgsqlDataReader datar = command.ExecuteReader();
            while (datar.Read())
            {
                clb.Items.Add(datar.GetString(0));
            }
            //command = new NpgsqlCommand("select \"name_tii\" from \"tab_chak\" where \"name_rp\" = '" + cb.SelectedItem.ToString() + "'", connection1);
            //datar = command.ExecuteReader();

            //while (datar.Read())
            //{
            //    listok.Add(datar.GetString(0));
            //}
            connection.Close();
            //connection1.Close();
            //if (listok.Count > 0)
            //{
            //    for (int i = 0; i < listok.Count; i++)
            //    {
            //        for (int j = 0; j < clb.Items.Count; j++)
            //        {
            //            if (listok[i].Equals(clb.Items[j].ToString()))
            //            {
            //                clb.SetItemChecked(j, true);
            //            }
            //        }
            //    }
            //}



            return clb;
        } //Загрузка списка чекбоксов 
        public CheckedListBox spisok_othod_rp_energo(CheckedListBox clb, ComboBox cb)
        {
            //List<string> listok = new List<string>();
            NpgsqlConnection connection = new NpgsqlConnection(energomera_comobjects);
            //NpgsqlConnection connection1 = new NpgsqlConnection(monitor_options);
            clb.Items.Clear();
            connection.Open();
            //connection1.Open();
            //NpgsqlCommand command = new NpgsqlCommand("select \"SerialNumber\" from \"Meters\" where \"ComObjectId\" IN (select \"ComObjectId\" from \"ComObjects\" where \"Name\" = '" + cb.SelectedItem.ToString() + "')", connection);
            NpgsqlCommand command = new NpgsqlCommand("select \"Name\" from \"BalanceGroups\" where \"BalanceGroupId\" IN (select \"BalanceGroupId\" from \"BalanceGroups_Meters\" where \"MeterId\" IN (SELECT \"MeterId\" FROM \"Meters\" where \"ComObjectId\" IN (select \"ComObjectId\" from \"ComObjects\" where \"Name\" = '" + cb.SelectedItem.ToString() + "')))", connection);
            NpgsqlDataReader datar = command.ExecuteReader();
            while (datar.Read()) { clb.Items.Add(datar.GetString(0)); }
            //command = new NpgsqlCommand("select \"name_tii\" from \"tab_chak\" where \"name_rp\" = '" + cb.SelectedItem.ToString() + "'", connection1);
            //datar = command.ExecuteReader();

            //while (datar.Read())
            //{
            //    listok.Add(datar.GetString(0));
            //}
            connection.Close();
            //connection1.Close();
            //if (listok.Count > 0)
            //{
            //    for (int i = 0; i < listok.Count; i++)
            //    {
            //        for (int j = 0; j < clb.Items.Count; j++)
            //        {
            //            if (listok[i].Equals(clb.Items[j].ToString()))
            //            {
            //                clb.SetItemChecked(j, true);
            //            }
            //        }
            //    }
            //}
            return clb;
        } //список отходящих для энергомеры
        public DataGridView resultat(DataGridView dgv, ComboBox cb_rp, CheckedListBox clb, DateTimePicker dtp_n, DateTimePicker dtp_k, CheckBox chb_energiya, ComboBox cb_vrem_n, ComboBox cb_vrem_k, ContextMenuStrip cms, RichTextBox rtb, ComboBox cb_base)
        {
            string z = "";
            dgv.DataSource = null;
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            result.Clear();
            cms.Items[4].Enabled = false;
            cms.Items[6].Enabled = false;
            cms.Items[7].Enabled = false;
            // rtb.Text = "Результат \n \n";
            NpgsqlConnection connection = new NpgsqlConnection(monitor);
            DataTable table = new DataTable();
            connection.Open();

            if (cb_base.SelectedItem.ToString() == "Компас РП") { z = "select t4.\"Values_ID\", t0.\"Name\", t4.\"Name\", t4.\"Tag\" from (select t1.\"Groups_ID\", t2.\"Name\" from \"GROUPS\" t1, \"GROUPS\" t2 where t1.\"Tag\" = 'ТИИ' and t1.\"Parent\"=t2.\"Groups_ID\" and t2.\"Name\" = '" + cb_rp.SelectedItem.ToString() + "') as t0, \"TB_VALUES\" t4 where t4.\"Groups_ID\"=t0.\"Groups_ID\" and t4.\"Name\" in  (" + tii(clb) + ")"; }
            if (cb_base.SelectedItem.ToString() == "Компас РП измер") { z = "select t4.\"Values_ID\", t0.\"Name\", t4.\"Name\", t4.\"Tag\" from (select t1.\"Groups_ID\", t1.\"Name\" from \"GROUPS\" t1 where t1.\"Name\" = '" + cb_rp.SelectedItem.ToString() + "') as t0, \"TB_VALUES\" t4 where t4.\"Groups_ID\"=t0.\"Groups_ID\" and t4.\"Name\" in  (" + tii(clb) + ")"; }
            //if (cb.SelectedItem == null) { MessageBox.Show("Выбрать РП"); return; }
            NpgsqlCommand com = new NpgsqlCommand(z, connection);
            rtb.Text = rtb.Text + "Результат \n \n" + z + "\n \n";
            NpgsqlDataAdapter dap = new NpgsqlDataAdapter(com);
            dap.Fill(table);
            connection.Close();
            list1.Clear();
            foreach (DataRow row in table.Rows)
            {
                class2 m = new class2();
                m.values_id = row.ItemArray[0].ToString();
                m.name = row.ItemArray[1].ToString();
                m.name_f = row.ItemArray[2].ToString();
                m.tag = row.ItemArray[3].ToString();
                list1.Add(m);
            }

            StringBuilder v_id = new StringBuilder();
            v_id.Append("(");
            for (int i = 0; i < list1.Count; i++)
            {

                v_id.Append(list1[i].values_id);
                if (i != list1.Count - 1)
                {
                    v_id.Append(", ");
                }
            }
            v_id.Append(")");

            DateTime d1 = new DateTime(dtp_n.Value.Year, dtp_n.Value.Month, dtp_n.Value.Day); //дата начала периода
            double x1 = d1.ToOADate();    //перевод нормальной даты в абсолютную

            DateTime d2 = new DateTime(dtp_k.Value.Year, dtp_k.Value.Month, dtp_k.Value.Day); //дата конца периода
            double x2 = d2.ToOADate();    //перевод нормальной даты в абсолютную

            long base_tick = new DateTime(1601, 1, 1).Ticks;
            long n_tick = d1.Ticks;
            long n_tick_mysql = n_tick-base_tick;
            long k_tick = d2.Ticks;
            long k_tick_mysql = k_tick - base_tick;

            //SqlConnection con = new SqlConnection(sql_connection);
            MySqlConnection con = new MySqlConnection(mysql_connection); //Для MariaDB
            con.Open();

            //if (x1 > x2) { MessageBox.Show("Неверный интервал времени."); }

            if (chb_energiya.Checked == true & cb_base.SelectedItem.ToString() == "Компас РП")
            {

                string rp = cb_rp.SelectedItem.ToString();
                List<string> list_f1 = new List<string>(); //лист фидеров

                if (x1 > x2) { MessageBox.Show("Неверный интервал времени."); }

                NpgsqlConnection connection1 = new NpgsqlConnection(monitor_options);
                connection1.Open();
                NpgsqlCommand com1 = new NpgsqlCommand("select \"name_tii\", \"ktt\"  from \"balans\" where \"name_rp\" = '" + rp + "' ", connection1);
                rtb.Text = rtb.Text + "select \"name_tii\", \"ktt\"  from \"balans\" where \"name_rp\" = '" + rp + "' \n \n";
                NpgsqlDataReader datar1 = com1.ExecuteReader();
                fider.Clear();
                while (datar1.Read())
                {
                    class5 m = new class5();
                    m.fider = datar1.GetValue(0).ToString();
                    m.val = datar1.GetValue(1).ToString();
                    fider.Add(m);

                }
                connection1.Close();
                schet.Clear();
                NpgsqlConnection connection2 = new NpgsqlConnection(monitor_options);
                connection2.Open();
                NpgsqlCommand com2 = new NpgsqlCommand("select \"name_tii\", \"num_sch\"  from \"number_sch\" where \"name_rp\" = '" + rp + "'", connection2);
                rtb.Text = rtb.Text + "select \"name_tii\", \"num_sch\"  from \"number_sch\" where \"name_rp\" = '" + rp + "' \n \n";
                NpgsqlDataReader datar2 = com2.ExecuteReader();
                //connection2.Close();

                while (datar2.Read())
                {
                    class6 m = new class6();
                    m.name_tii = datar2.GetValue(0).ToString();
                    m.num_schet = datar2.GetValue(1).ToString();
                    schet.Add(m);

                }
                connection2.Close();
                result.Clear();



                list_sutki.Clear();

                //byte dayCount = Convert.ToByte(x2 - x1);
                ////MessageBox.Show(dayCount.ToString());
                //for (byte i = 0; i < dayCount + 1; i++)
                //{
                //    for (int j = 0; j < list1.Count; j++)
                //    {
                //        //SqlCommand com10 = new SqlCommand("SELECT max(V_VALUE) - min(V_VALUE) FROM ti  where VALUES_ID = " + list1[j].values_id + " and T_TIME >= " + (x1 + i) + "and T_TIME <= " + (x1 + i + 1), con);
                //        MySqlCommand com10 = new MySqlCommand("SELECT max(V_VALUE) - min(V_VALUE) FROM ti  where VALUES_ID = " + list1[j].values_id + " and T_TIME >= " + (x1 + i) + "and T_TIME <= " + (x1 + i + 1), con); //Для MariaDB
                //        rtb.Text = rtb.Text + "SELECT max(V_VALUE) - min(V_VALUE) FROM ti  where VALUES_ID = " + list1[j].values_id + " and T_TIME >= " + (x1 + i) + "and T_TIME <= " + (x1 + i + 1) + "\n \n";
                //        com10.ExecuteNonQuery();
                //        Decimal s1 = Convert.ToDecimal(com10.ExecuteScalar());
                //        class_sutki m = new class_sutki();
                //        m.id = list1[j].values_id;
                //        m.name = list1[j].name_f.ToString();
                //        m.value = s1;
                //        m.data = i;
                //        list_sutki.Add(m);
                //    }
                //}


                ////-----------------Для MariaDB-----------------------------------
                long dayCount = k_tick_mysql - n_tick_mysql;
                long oneday = 864000000000;
                //MessageBox.Show(dayCount.ToString());
                for (long i = 0; i < dayCount + oneday; i=i+oneday)
                {
                    for (int j = 0; j < list1.Count; j++)
                    {
                        MySqlCommand com10 = new MySqlCommand("SELECT max(V_VALUE) - min(V_VALUE) FROM ti  where VALUES_ID = " + list1[j].values_id + " and T_TIME >= " + (n_tick_mysql + i) + " and T_TIME <= " + (n_tick_mysql + i + oneday), con); //Для MariaDB
                        rtb.Text = rtb.Text + "SELECT max(V_VALUE) - min(V_VALUE) FROM ti  where VALUES_ID = " + list1[j].values_id + " and T_TIME >= " + (n_tick_mysql + i) + " and T_TIME <= " + (n_tick_mysql + i + oneday) + "\n \n";
                        Decimal s1;
                        try
                        {
                            com10.ExecuteNonQuery();
                             s1 = Convert.ToDecimal(com10.ExecuteScalar());
                        }
                        catch { MessageBox.Show("Нет данных за период " + new DateTime(n_tick_mysql + i + base_tick) + "-" + new DateTime(n_tick_mysql + i + oneday + base_tick)); break; }
                        class_sutki m = new class_sutki();
                        m.id = list1[j].values_id;
                        m.name = list1[j].name_f.ToString();
                        m.value = s1;
                        m.data = i;
                        list_sutki.Add(m);
                        
                    }
                }
                ////-----------------Для MariaDB-----------------------------------


                for (int i = 0; i < list_sutki.Count; i++)
                {
                    for (int j = 0; j < list1.Count; j++)
                    {
                        if (list_sutki[i].id == list1[j].values_id)
                        {

                            class3 m = new class3();
                            m.name = list1[j].name;
                            m.name_f = list1[j].name_f;
                            m.tag = list1[j].tag;
                           // m.t_time = DateTime.FromOADate((x1 + list_sutki[i].data)).ToString();
                            m.t_time = (new DateTime(n_tick_mysql+list_sutki[i].data + base_tick)).ToString(); //Для MariaDB
                            m.v_value = Math.Round(list_sutki[i].value, 3).ToString();


                            for (int k = 0; k < fider.Count; k++)
                            {

                                if (list1[j].name_f.Equals(fider[k].fider))
                                { m.ktt = fider[k].val; }


                            }
                            for (int l = 0; l < schet.Count; l++)
                            {
                                if (list1[j].name_f.Equals(schet[l].name_tii))
                                {

                                    m.schet = schet[l].num_schet;
                                }
                            }
                            result.Add(m);
                        }
                    }
                }

                //var sum = result.Sum(p => Convert.ToDecimal(p.v_value));
                class3 m1 = new class3();
                m1.name = "Итог";
                m1.t_time = "За период с " + d1 + " до " + d2;
                m1.name_f = "";
                result.Add(m1);

                dgv.DataSource = result;

                for (int i = 0; i < dgv.Rows.Count; i++)
                {

                    dgv.Rows[i].Cells[6].Value = (Convert.ToDecimal(dgv.Rows[i].Cells[6].Value) * Convert.ToDecimal(dgv.Rows[i].Cells[4].Value)).ToString();

                }

                dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells); //автоматический размер колонок
                dgv.Columns[2].Visible = false;  //скрываем столбец с названием ТИИ
                dgv.Columns[7].Visible = false;
                dgv.Columns[0].HeaderText = "Объект";
                dgv.Columns[1].HeaderText = "Присоединение"; //присваиваем текст к заголовку колонок
                dgv.Columns[3].HeaderText = "Счетчик №";
                dgv.Columns[4].HeaderText = "Кт";
                dgv.Columns[5].HeaderText = "Дата и время";
                dgv.Columns[6].HeaderText = "Энергия(кВт*ч)";
                //dataGridView3.Columns[7].HeaderText = "*";
                // dgv.RowHeadersWidth = 20;


                List<string> s_list = new List<string>();
                List<string[]> sum_list = new List<string[]>();
                s_list.Clear();
                sum_list.Clear();
                string stroka = "";
                string strochka = "";
                string strochka1 = "";
                for (int i = 0; i < result.Count - 1; i++)
                {

                    if (!s_list.Contains(result[i].name_f)) { s_list.Add(result[i].name_f); }
                }

                for (int k = 0; k < s_list.Count; k++)
                {
                    string[] str = { s_list[k], "0" };
                    sum_list.Add(str);
                }

                for (int k = 0; k < sum_list.Count; k++)
                {
                    for (int i = 0; i < result.Count; i++)
                    {
                        if (sum_list[k][0].Equals(result[i].name_f))
                        {
                            sum_list[k][1] = Math.Round((Convert.ToDecimal(sum_list[k][1]) + Convert.ToDecimal(result[i].v_value)), 3).ToString();
                        }

                    }
                }


                for (int j = 0; j < sum_list.Count; j++)
                {
                    stroka = stroka + sum_list[j][0].ToString() + " = " + sum_list[j][1].ToString() + "\n";
                    strochka1 = strochka1 + sum_list[j][0].ToString() + "\n";
                    strochka = strochka + sum_list[j][1].ToString() + "\n";

                }

                int nd = DateTime.FromOADate(x1).ToString().IndexOf(" ");
                int kd = DateTime.FromOADate(x2).ToString().IndexOf(" ");






                for (int i = 0; i < dgv.Rows.Count; i++)
                {
                    //dataGridView3.Rows.Add();

                    //dgv.Rows[dgv.Rows.Count - 1].Cells[5].Value = "ИТОГО";
                    dgv.Rows[dgv.Rows.Count - 1].Cells[6].Value = "\n" + stroka;
                }
                //MessageBox.Show(strochka);

                //MessageBox.Show("Энергия с " + DateTime.FromOADate(x1).ToString().Substring(0, nd) + " по " + DateTime.FromOADate(x2).ToString().Substring(0, kd) + "\n" + stroka, "Энергия");
            }

            if (chb_energiya.Checked == false & cb_base.SelectedItem.ToString() == "Компас РП")
            {

                string chas1 = cb_vrem_n.SelectedItem.ToString();
                chas1 = chas1.Substring(0, chas1.IndexOf(":"));
                string min1 = cb_vrem_n.SelectedItem.ToString();
                min1 = min1.Substring(min1.IndexOf(":") + 1, 2);
                double ab_vr1 = (Convert.ToDouble(chas1) * 60 + Convert.ToDouble(min1)) * 0.000694443; //перевод часов и минут (то есть время) в абсолютное. Это от
                long tick_vr1 = Convert.ToInt64(chas1) * 36000000000 + Convert.ToInt64(min1) * 600000000; //Для mariadb


                string chas2 = cb_vrem_k.SelectedItem.ToString();
                chas2 = chas2.Substring(0, chas2.IndexOf(":"));
                string min2 = cb_vrem_k.SelectedItem.ToString();
                min2 = min2.Substring(min2.IndexOf(":") + 1, 2);
                double ab_vr2 = (Convert.ToDouble(chas2) * 60 + Convert.ToDouble(min2)) * 0.000694445;  //перевод часов и минут (то есть время) в абсолютное. Это до
                long tick_vr2 = Convert.ToInt64(chas2) * 36000000000 + Convert.ToInt64(min2) * 600000000; //Для mariadb

                //double vr1 = x1 + ab_vr1;
                //double vr2 = x2 + ab_vr2;
                //string vr11 = Convert.ToString(vr1).Replace(",", ".");
                //string vr22 = Convert.ToString(vr2).Replace(",", ".");

                long vr1 = n_tick_mysql + tick_vr1; //Для mariadb
                long vr2 = k_tick_mysql + tick_vr2; //Для mariadb

                if (vr1 > vr2) { MessageBox.Show("Неверный интервал времени."); }

                
                

                //SqlDataAdapter da = new SqlDataAdapter("select VALUES_ID, T_TIME,  V_VALUE, VALID from ti where T_TIME >= " + vr11 + " and T_TIME <= " + vr22 + " and values_id in " + v_id.ToString() + " order by VALUES_ID, T_TIME ", con);
                //MySqlDataAdapter da = new MySqlDataAdapter("select VALUES_ID, T_TIME,  V_VALUE, VALID from ti where T_TIME >= " + vr11 + " and T_TIME <= " + vr22 + " and values_id in " + v_id.ToString() + " order by VALUES_ID, T_TIME ", con);
                //rtb.Text = rtb.Text + "select VALUES_ID, T_TIME,  V_VALUE, VALID from ti where T_TIME >= " + vr11 + " and T_TIME <= " + vr22 + " and values_id in " + v_id.ToString() + " order by VALUES_ID, T_TIME " + "\n \n";
                MySqlDataAdapter da = new MySqlDataAdapter("select VALUES_ID, T_TIME,  V_VALUE, VALID from ti where T_TIME >= " + vr1 + " and T_TIME <= " + vr2 + " and values_id in " + v_id.ToString() + " order by VALUES_ID, T_TIME ", con);
                rtb.Text = rtb.Text + "select VALUES_ID, T_TIME,  V_VALUE, VALID from ti where T_TIME >= " + vr1 + " and T_TIME <= " + vr2 + " and values_id in " + v_id.ToString() + " order by VALUES_ID, T_TIME " + "\n \n";
                //SqlCommandBuilder cb = new SqlCommandBuilder(da);
                MySqlCommandBuilder cb = new MySqlCommandBuilder(da);
                DataTable table1 = new DataTable();
                da.Fill(table1);
                list.Clear();
                foreach (DataRow row in table1.Rows)
                {
                    class0 m = new class0();
                    m.t_time = row.ItemArray[1].ToString();
                    m.values_id = row.ItemArray[0].ToString();
                    m.v_value = row.ItemArray[2].ToString();
                    m.valid = row.ItemArray[3].ToString();
                    list.Add(m);
                }
                con.Close();
                //for (int i = 0; i < list.Count; i++)
                //{
                //    double x = Convert.ToDouble(list[i].t_time);
                //    DateTime dt = DateTime.FromOADate(x);
                //    list[i].t_time = dt.ToString();
                //}
                
                //---------------для MariaDB--------------
                for (int i = 0; i < list.Count; i++)
                {
                    long x = Convert.ToInt64(list[i].t_time)+base_tick;
                    DateTime dt = new DateTime(x);
                    list[i].t_time = dt.ToString();
                }
                //---------------для MariaDB--------------


                string rp = cb_rp.SelectedItem.ToString();
                List<string> list_f1 = new List<string>(); //лист фидеров
                NpgsqlConnection connection1 = new NpgsqlConnection(monitor_options);
                connection1.Open();
                NpgsqlCommand com1 = new NpgsqlCommand("select \"name_tii\", \"ktt\"  from \"balans\" where \"name_rp\" = '" + rp + "' ", connection1);
                rtb.Text = rtb.Text + "select \"name_tii\", \"ktt\"  from \"balans\" where \"name_rp\" = '" + rp + "' \n \n";
                NpgsqlDataReader datar1 = com1.ExecuteReader();
                fider.Clear();
                while (datar1.Read())
                {
                    class5 m = new class5();
                    m.fider = datar1.GetValue(0).ToString();
                    m.val = datar1.GetValue(1).ToString();
                    fider.Add(m);

                }
                connection1.Close();
                schet.Clear();
                NpgsqlConnection connection2 = new NpgsqlConnection(monitor_options);
                connection2.Open();
                NpgsqlCommand com2 = new NpgsqlCommand("select \"name_tii\", \"num_sch\"  from \"number_sch\" where \"name_rp\" = '" + rp + "'", connection2);
                rtb.Text = rtb.Text + "select \"name_tii\", \"num_sch\"  from \"number_sch\" where \"name_rp\" = '" + rp + "' \n \n";
                NpgsqlDataReader datar2 = com2.ExecuteReader();
                //connection2.Close();

                while (datar2.Read())
                {
                    class6 m = new class6();
                    m.name_tii = datar2.GetValue(0).ToString();
                    m.num_schet = datar2.GetValue(1).ToString();
                    schet.Add(m);

                }
                connection2.Close();
                result.Clear();
                for (int i = 0; i < list.Count; i++)
                {
                    for (int j = 0; j < list1.Count; j++)
                    {
                        if (list[i].values_id == list1[j].values_id)
                        {
                            class3 m = new class3();
                            m.name = list1[j].name;
                            m.name_f = list1[j].name_f;
                            m.tag = list1[j].tag;
                            m.t_time = list[i].t_time;
                            m.v_value = list[i].v_value;
                            m.valid = list[i].valid;
                            //m.meas = list1[j].meas;
                            result.Add(m);
                        }
                    }
                }
                dgv.DataSource = null;  //обнуляем таблицу

                dgv.DataSource = result; //заполняем таблицу листом result
                for (int i = 0; i < dgv.Rows.Count; i++)
                {
                    for (int j = 0; j < dgv.Columns.Count; j++)
                    {
                        for (int k = 0; k < fider.Count; k++)
                        {
                            if (dgv.Rows[i].Cells[1].Value.ToString().Equals(fider[k].fider)) { dgv.Rows[i].Cells[4].Value = fider[k].val; }
                        }
                        for (int n = 0; n < schet.Count; n++)
                        {
                            string sk = dgv.Rows[i].Cells[1].Value.ToString();
                            string sn = schet[n].name_tii;
                            // if (dgv.Rows[i].Cells[1].Value.ToString().Equals(schet[n].name_tii)) { dgv.Rows[i].Cells[3].Value = schet[n].num_schet; }
                            if (sk.Substring(sk.IndexOf(' '), sk.Length - sk.IndexOf(' ')).Equals(sn.Substring(sn.IndexOf(' '), sn.Length - sn.IndexOf(' ')))) { dgv.Rows[i].Cells[3].Value = schet[n].num_schet; }
                        }
                    }
                }

                dgv.Columns[2].Visible = false;  //скрываем столбец с названием ТИИ
                dgv.Columns[0].HeaderText = "Объект";
                dgv.Columns[1].HeaderText = "Присоединение"; //присваиваем текст к заголовку колонок
                dgv.Columns[3].HeaderText = "Счетчик №";
                dgv.Columns[4].HeaderText = "Кт";
                dgv.Columns[5].HeaderText = "Дата и время";
                dgv.Columns[6].HeaderText = "кВт*ч";
                dgv.Columns[7].HeaderText = "*";

            }

            if (cb_base.SelectedItem.ToString() == "Компас РП измер")
            {

                string chas1 = cb_vrem_n.SelectedItem.ToString();
                chas1 = chas1.Substring(0, chas1.IndexOf(":"));
                string min1 = cb_vrem_n.SelectedItem.ToString();
                min1 = min1.Substring(min1.IndexOf(":") + 1, 2);
                double ab_vr1 = (Convert.ToDouble(chas1) * 60 + Convert.ToDouble(min1)) * 0.000694443; //перевод часов и минут (то есть время) в абсолютное. Это от


                string chas2 = cb_vrem_k.SelectedItem.ToString();
                chas2 = chas2.Substring(0, chas2.IndexOf(":"));
                string min2 = cb_vrem_k.SelectedItem.ToString();
                min2 = min2.Substring(min2.IndexOf(":") + 1, 2);
                double ab_vr2 = (Convert.ToDouble(chas2) * 60 + Convert.ToDouble(min2)) * 0.000694445;  //перевод часов и минут (то есть время) в абсолютное. Это до

                //double vr1 = x1 + ab_vr1;
                //double vr2 = x2 + ab_vr2;
                //string vr11 = Convert.ToString(vr1).Replace(",", ".");
                //string vr22 = Convert.ToString(vr2).Replace(",", ".");

                long tick_vr1 = Convert.ToInt64(chas1) * 36000000000 + Convert.ToInt64(min1) * 600000000; //Для mariadb
                long tick_vr2 = Convert.ToInt64(chas2) * 36000000000 + Convert.ToInt64(min2) * 600000000; //Для mariadb

                long vr1 = n_tick_mysql + tick_vr1; //Для mariadb
                long vr2 = k_tick_mysql + tick_vr2; //Для mariadb

                if (vr1 > vr2) { MessageBox.Show("Неверный интервал времени."); }

                //SqlDataAdapter da = new SqlDataAdapter("select VALUES_ID, T_TIME,  V_VALUE, VALID from ti where T_TIME >= " + vr11 + " and T_TIME <= " + vr22 + " and values_id in " + v_id.ToString() + " order by VALUES_ID, T_TIME ", con);
                //MySqlDataAdapter da = new MySqlDataAdapter("select VALUES_ID, T_TIME,  V_VALUE, VALID from ti where T_TIME >= " + vr11 + " and T_TIME <= " + vr22 + " and values_id in " + v_id.ToString() + " order by VALUES_ID, T_TIME ", con);
                //rtb.Text = rtb.Text + "select VALUES_ID, T_TIME,  V_VALUE, VALID from ti where T_TIME >= " + vr11 + " and T_TIME <= " + vr22 + " and values_id in " + v_id.ToString() + " order by VALUES_ID, T_TIME " + "\n \n";
                MySqlDataAdapter da = new MySqlDataAdapter("select VALUES_ID, T_TIME,  V_VALUE, VALID from ti where T_TIME >= " + vr1 + " and T_TIME <= " + vr2 + " and values_id in " + v_id.ToString() + " order by VALUES_ID, T_TIME ", con);
                rtb.Text = rtb.Text + "select VALUES_ID, T_TIME,  V_VALUE, VALID from ti where T_TIME >= " + vr1 + " and T_TIME <= " + vr2 + " and values_id in " + v_id.ToString() + " order by VALUES_ID, T_TIME " + "\n \n";
                //SqlCommandBuilder cb = new SqlCommandBuilder(da);
                MySqlCommandBuilder cb = new MySqlCommandBuilder(da);
                DataTable table1 = new DataTable();
                da.Fill(table1);
                list.Clear();
                foreach (DataRow row in table1.Rows)
                {
                    class0 m = new class0();
                    m.t_time = row.ItemArray[1].ToString();
                    m.values_id = row.ItemArray[0].ToString();
                    m.v_value = row.ItemArray[2].ToString();
                    m.valid = row.ItemArray[3].ToString();
                    list.Add(m);
                }
                con.Close();

                //for (int i = 0; i < list.Count; i++)
                //{
                //    double x = Convert.ToDouble(list[i].t_time);
                //    DateTime dt = DateTime.FromOADate(x);
                //    list[i].t_time = dt.ToString();
                //}


                //---------------для MariaDB--------------
                for (int i = 0; i < list.Count; i++)
                {
                    long x = Convert.ToInt64(list[i].t_time)+base_tick;
                    DateTime dt = new DateTime(x);
                    list[i].t_time = dt.ToString();
                }
                //---------------для MariaDB--------------

                string rp = cb_rp.SelectedItem.ToString();
                List<string> list_f1 = new List<string>(); //лист фидеров
                NpgsqlConnection connection1 = new NpgsqlConnection(monitor_options);
                connection1.Open();
                NpgsqlCommand com1 = new NpgsqlCommand("select \"name_tii\", \"ktt\"  from \"balans\" where \"name_rp\" = '" + rp + "' ", connection1);
                rtb.Text = rtb.Text + "select \"name_tii\", \"ktt\"  from \"balans\" where \"name_rp\" = '" + rp + "' \n \n";
                NpgsqlDataReader datar1 = com1.ExecuteReader();
                fider.Clear();
                while (datar1.Read())
                {
                    class5 m = new class5();
                    m.fider = datar1.GetValue(0).ToString();
                    m.val = datar1.GetValue(1).ToString();
                    fider.Add(m);

                }
                connection1.Close();
                schet.Clear();
                NpgsqlConnection connection2 = new NpgsqlConnection(monitor_options);
                connection2.Open();
                NpgsqlCommand com2 = new NpgsqlCommand("select \"name_tii\", \"num_sch\"  from \"number_sch\" where \"name_rp\" = '" + rp + "'", connection2);
                rtb.Text = rtb.Text + "select \"name_tii\", \"num_sch\"  from \"number_sch\" where \"name_rp\" = '" + rp + "' \n \n";
                NpgsqlDataReader datar2 = com2.ExecuteReader();
                //connection2.Close();

                while (datar2.Read())
                {
                    class6 m = new class6();
                    m.name_tii = datar2.GetValue(0).ToString();
                    m.num_schet = datar2.GetValue(1).ToString();
                    schet.Add(m);

                }
                connection2.Close();

                result.Clear();

                for (int i = 0; i < list.Count; i++)
                {
                    for (int j = 0; j < list1.Count; j++)
                    {

                        if (list[i].values_id == list1[j].values_id)
                        {
                            class3 m = new class3();
                            m.name = list1[j].name;
                            m.name_f = list1[j].name_f;
                            m.tag = list1[j].tag;
                            m.t_time = list[i].t_time;
                            m.v_value = list[i].v_value;
                            m.valid = list[i].valid;
                            //m.meas = list1[j].meas;
                            result.Add(m);

                        }
                    }
                }
                dgv.DataSource = null;  //обнуляем таблицу
                dgv.DataSource = result; //заполняем таблицу листом result

                for (int i = 0; i < dgv.Rows.Count; i++)
                {
                    for (int j = 0; j < dgv.Columns.Count; j++)
                    {
                        for (int k = 0; k < fider.Count; k++)
                        {
                            if (dgv.Rows[i].Cells[1].Value.ToString().Equals(fider[k].fider)) { dgv.Rows[i].Cells[4].Value = fider[k].val; }
                        }

                        for (int n = 0; n < schet.Count; n++)
                        {
                            string sk = dgv.Rows[i].Cells[1].Value.ToString();
                            string sn = schet[n].name_tii;
                            // if (dgv.Rows[i].Cells[1].Value.ToString().Equals(schet[n].name_tii)) { dgv.Rows[i].Cells[3].Value = schet[n].num_schet; }
                            if (sk.Substring(sk.IndexOf(' '), sk.Length - sk.IndexOf(' ')).Equals(sn.Substring(sn.IndexOf(' '), sn.Length - sn.IndexOf(' ')))) { dgv.Rows[i].Cells[3].Value = schet[n].num_schet; }
                        }



                    }
                }

                dgv.Columns[2].Visible = false;  //скрываем столбец с названием ТИИ
                dgv.Columns[0].HeaderText = "Объект";
                dgv.Columns[1].HeaderText = "Присоединение"; //присваиваем текст к заголовку колонок
                dgv.Columns[3].HeaderText = "Счетчик №";
                dgv.Columns[4].HeaderText = "Кт";
                dgv.Columns[5].HeaderText = "Дата и время";
                dgv.Columns[6].HeaderText = "Значение";
                dgv.Columns[7].HeaderText = "*";


            }

            vrem_tabl.Clear();
            finish.Clear();

            string s;
            for (int i = 0; i < result.Count; i++)
            {
                class_vrem_tabl m = new class_vrem_tabl();
                m.id = i;

                s = result[i].name_f;
                if (s.Contains("Вв1") || s.Contains("вв1") || s.Contains("Вв 1") || s.Contains("вв 1"))
                {
                    m.type = 1;
                    m.num = 0;
                }
                else if (s.Contains("Вв2") || s.Contains("вв2") || s.Contains("Вв 2") || s.Contains("вв 2"))
                {
                    m.type = 3;
                    m.num = 0;
                }

                else if (s.Contains("ф") & !s.Contains("обр"))
                {
                    string str = s.Substring(s.IndexOf("ф") + 1, (s.Length - s.IndexOf("ф")) - 1);
                    int nomer = Convert.ToInt32(str);
                    if (nomer % 2 != 0)
                    {
                        m.type = 2;
                    }

                    else { m.type = 4; }
                    m.num = nomer;
                }
                else if (s.Contains("Т-1") || s.Contains("Т-2") || s.Contains("обр"))
                {
                    m.type = 5;
                    m.num = 0;
                }
                else
                {
                    m.type = 6;
                    m.num = 0;
                }
                vrem_tabl.Add(m);

            }
            IList<class_vrem_tabl> res = vrem_tabl.OrderBy(x => x.type).ThenBy(x => x.num).ToList(); //сортировка

            for (int i = 0; i < res.Count; i++)
            {
                class_finish m = new class_finish();
                m.name = result[res[i].id].name;
                m.name_f = result[res[i].id].name_f;
                m.tag = result[res[i].id].tag;
                m.schet = result[res[i].id].schet;
                m.ktt = result[res[i].id].ktt;
                m.t_time = result[res[i].id].t_time;
                m.v_value = result[res[i].id].v_value;
                m.valid = result[res[i].id].valid;
                finish.Add(m);
            }


            dgv.DataSource = finish;
            dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                for (int j = 0; j < dgv.Columns.Count; j++)
                {
                    try
                    {
                        if (dgv.Rows[i].Cells[7].Value.ToString() == "0") { dgv.Rows[i].DefaultCellStyle.BackColor = Color.Red; }
                    }
                    catch { }

                }
            }

            return dgv;



        } //Кнопка результат, вывод таблицы с данными
        public DataGridView resultat_energomera(DataGridView dgv, ComboBox cb_rp, CheckedListBox clb, DateTimePicker dtp_n, DateTimePicker dtp_k, ComboBox cb_vrem_n, ComboBox cb_vrem_k, ContextMenuStrip cms, RichTextBox rtb, CheckBox chb_energomera)
        {
            dgv.DataSource = null;
            dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            vivod_energomera.Clear();

            DateTime d1 = new DateTime(dtp_n.Value.Year, dtp_n.Value.Month, dtp_n.Value.Day); //дата начала периода
            double x1 = d1.ToOADate();    //перевод нормальной даты в абсолютную

            DateTime d2 = new DateTime(dtp_k.Value.Year, dtp_k.Value.Month, dtp_k.Value.Day); //дата конца периода
            double x2 = d2.ToOADate();    //перевод нормальной даты в абсолютную

            NpgsqlConnection conn = new NpgsqlConnection(energomera_comobjects);
            DataTable table11 = new DataTable();
            conn.Open();
            //NpgsqlCommand com = new NpgsqlCommand(" SELECT \"Meters\".\"SerialNumber\",\"Values\".\"DT\", \"Values\".\"Val\"  FROM \"Values\", \"Meters\"  where \"Values\".\"MeterId\" IN (SELECT \"MeterId\" FROM \"Meters\" where \"ComObjectId\" IN (select \"ComObjectId\" from \"ComObjects\" where \"Name\" = '" + cb_rp.SelectedItem.ToString() + "')) and \"Values\".\"MeterId\"=\"Meters\".\"MeterId\" and \"Values\".\"TariffId\" = 1 and \"Meters\".\"SerialNumber\" IN (" + tii(clb) + ") and \"Values\".\"DT\" >= '" + n_d + " ' and \"Values\".\"DT\" <= '" + k_d +" ' ", connection);
            NpgsqlCommand com11 = new NpgsqlCommand(" select  balancegroup.\"Name\", valuesmeters.\"SerialNumber\", valuesmeters.\"MeterId\" from (SELECT  \"Meters\".\"SerialNumber\", \"Meters\".\"MeterId\" from \"Meters\" where \"Meters\".\"MeterId\" IN (select \"MeterId\" from \"BalanceGroups_Meters\" where \"BalanceGroupId\" IN (select \"BalanceGroupId\" from \"BalanceGroups\" where \"Name\" IN (" + tii(clb) + "))) and \"Meters\".\"SerialNumber\" IN (select \"SerialNumber\" from \"Meters\" where \"ComObjectId\" IN (select \"ComObjectId\" from \"ComObjects\" where \"Name\" = '" + cb_rp.SelectedItem.ToString() + "'))) as valuesmeters left join (SELECT  \"BalanceGroups_Meters\".\"BalanceGroupId\", \"BalanceGroups_Meters\".\"MeterId\", \"BalanceGroups\".\"Name\" FROM \"BalanceGroups_Meters\", \"BalanceGroups\" WHERE \"BalanceGroups\".\"BalanceGroupId\"=\"BalanceGroups_Meters\".\"BalanceGroupId\") as balancegroup on balancegroup.\"MeterId\"=valuesmeters.\"MeterId\" ", conn);
            //MessageBox.Show(" select  balancegroup.\"Name\", valuesmeters.\"SerialNumber\", valuesmeters.\"MeterId\" from (SELECT  \"Meters\".\"SerialNumber\", \"Meters\".\"MeterId\",  \"Values\".\"DT\", \"Values\".\"Val\" from  \"Values\", \"Meters\" where \"Values\".\"MeterId\" IN (select \"MeterId\" from \"BalanceGroups_Meters\" where \"BalanceGroupId\" IN (select \"BalanceGroupId\" from \"BalanceGroups\" where \"Name\" IN (" + tii(clb) + "))) and \"Values\".\"MeterId\"=\"Meters\".\"MeterId\" and \"Values\".\"TariffId\" = 1 and \"Meters\".\"SerialNumber\" IN (select \"SerialNumber\" from \"Meters\" where \"ComObjectId\" IN (select \"ComObjectId\" from \"ComObjects\" where \"Name\" = '" + cb_rp.SelectedItem.ToString() + "'))) as valuesmeters left join (SELECT  \"BalanceGroups_Meters\".\"BalanceGroupId\", \"BalanceGroups_Meters\".\"MeterId\", \"BalanceGroups\".\"Name\" FROM \"BalanceGroups_Meters\", \"BalanceGroups\" WHERE \"BalanceGroups\".\"BalanceGroupId\"=\"BalanceGroups_Meters\".\"BalanceGroupId\") as balancegroup on balancegroup.\"MeterId\"=valuesmeters.\"MeterId\" ");
            rtb.Text = rtb.Text + " select  balancegroup.\"Name\", valuesmeters.\"SerialNumber\", valuesmeters.\"MeterId\" from (SELECT  \"Meters\".\"SerialNumber\", \"Meters\".\"MeterId\" from \"Meters\" where \"Meters\".\"MeterId\" IN (select \"MeterId\" from \"BalanceGroups_Meters\" where \"BalanceGroupId\" IN (select \"BalanceGroupId\" from \"BalanceGroups\" where \"Name\" IN (" + tii(clb) + "))) and \"Meters\".\"SerialNumber\" IN (select \"SerialNumber\" from \"Meters\" where \"ComObjectId\" IN (select \"ComObjectId\" from \"ComObjects\" where \"Name\" = '" + cb_rp.SelectedItem.ToString() + "'))) as valuesmeters left join (SELECT  \"BalanceGroups_Meters\".\"BalanceGroupId\", \"BalanceGroups_Meters\".\"MeterId\", \"BalanceGroups\".\"Name\" FROM \"BalanceGroups_Meters\", \"BalanceGroups\" WHERE \"BalanceGroups\".\"BalanceGroupId\"=\"BalanceGroups_Meters\".\"BalanceGroupId\") as balancegroup on balancegroup.\"MeterId\"=valuesmeters.\"MeterId\" ";
            NpgsqlDataAdapter dap11 = new NpgsqlDataAdapter(com11);
            dap11.Fill(table11);
            conn.Close();

            list1_energo.Clear();
            foreach (DataRow row in table11.Rows)
            {
                class2_energo m = new class2_energo();
                m.meters_id = row.ItemArray[2].ToString();
                m.name = cb_rp.SelectedItem.ToString();
                m.name_f = row.ItemArray[0].ToString();
                m.sn = row.ItemArray[1].ToString();
                list1_energo.Add(m);
            }

            StringBuilder v_id = new StringBuilder();
            v_id.Append("(");
            for (int i = 0; i < list1_energo.Count; i++)
            {

                v_id.Append(list1_energo[i].meters_id);
                if (i != list1_energo.Count - 1)
                {
                    v_id.Append(", ");
                }
            }
            v_id.Append(")");

            if (chb_energomera.Checked == false)
            {

                NpgsqlConnection connection1 = new NpgsqlConnection(monitor_options);
                connection1.Open();
                NpgsqlCommand com1 = new NpgsqlCommand("select \"name_tii\", \"ktt\"  from \"balans\" where \"name_rp\" = '" + cb_rp.SelectedItem.ToString() + "' ", connection1);
                rtb.Text = rtb.Text + "select \"name_tii\", \"ktt\"  from \"balans\" where \"name_rp\" = '" + cb_rp.SelectedItem.ToString() + "' \n \n";
                NpgsqlDataReader datar1 = com1.ExecuteReader();
                energo_ktt.Clear();
                while (datar1.Read())
                {
                    class5_energo m = new class5_energo();
                    m.fider = datar1.GetValue(0).ToString();
                    m.val = datar1.GetValue(1).ToString();
                    energo_ktt.Add(m);
                }
                connection1.Close();


                string chas1 = cb_vrem_n.SelectedItem.ToString();
                chas1 = chas1.Substring(0, chas1.IndexOf(":"));
                string min1 = cb_vrem_n.SelectedItem.ToString();
                min1 = min1.Substring(min1.IndexOf(":") + 1, 2);
                double ab_vr1 = (Convert.ToDouble(chas1) * 60 + Convert.ToDouble(min1)) * 0.000694443; //перевод часов и минут (то есть время) в абсолютное. Это от


                string chas2 = cb_vrem_k.SelectedItem.ToString();
                chas2 = chas2.Substring(0, chas2.IndexOf(":"));
                string min2 = cb_vrem_k.SelectedItem.ToString();
                min2 = min2.Substring(min2.IndexOf(":") + 1, 2);
                double ab_vr2 = (Convert.ToDouble(chas2) * 60 + Convert.ToDouble(min2)) * 0.000694445;  //перевод часов и минут (то есть время) в абсолютное. Это до

                double vr1 = x1 + ab_vr1;
                double vr2 = x2 + ab_vr2;


                string n_d = DateTime.FromOADate(vr1).ToString().Replace(".", "-");
                string k_d = DateTime.FromOADate(vr2).ToString().Replace(".", "-");

                NpgsqlConnection connection = new NpgsqlConnection(energomera_comobjects);
                DataTable table = new DataTable();
                connection.Open();
                //NpgsqlCommand com = new NpgsqlCommand(" SELECT \"Meters\".\"SerialNumber\",\"Values\".\"DT\", \"Values\".\"Val\"  FROM \"Values\", \"Meters\"  where \"Values\".\"MeterId\" IN (SELECT \"MeterId\" FROM \"Meters\" where \"ComObjectId\" IN (select \"ComObjectId\" from \"ComObjects\" where \"Name\" = '" + cb_rp.SelectedItem.ToString() + "')) and \"Values\".\"MeterId\"=\"Meters\".\"MeterId\" and \"Values\".\"TariffId\" = 1 and \"Meters\".\"SerialNumber\" IN (" + tii(clb) + ") and \"Values\".\"DT\" >= '" + n_d + " ' and \"Values\".\"DT\" <= '" + k_d +" ' ", connection);
                NpgsqlCommand com = new NpgsqlCommand(" select  balancegroup.\"Name\", valuesmeters.\"SerialNumber\", valuesmeters.\"DT\", valuesmeters.\"Val\" from (SELECT  \"Meters\".\"SerialNumber\", \"Meters\".\"MeterId\",  \"ValueProfiles\".\"DT\", \"ValueProfiles\".\"Val\" from  \"ValueProfiles\", \"Meters\" where \"ValueProfiles\".\"MeterId\" IN (select \"MeterId\" from \"BalanceGroups_Meters\" where \"BalanceGroupId\" IN (select \"BalanceGroupId\" from \"BalanceGroups\" where \"Name\" IN (" + tii(clb) + "))) and \"ValueProfiles\".\"MeterId\"=\"Meters\".\"MeterId\" and \"ValueProfiles\".\"StateId\" = 1 and \"Meters\".\"SerialNumber\" IN (select \"SerialNumber\" from \"Meters\" where \"ComObjectId\" IN (select \"ComObjectId\" from \"ComObjects\" where \"Name\" = '" + cb_rp.SelectedItem.ToString() + "')) and \"ValueProfiles\".\"DT\" >= '" + n_d + "' and \"ValueProfiles\".\"DT\" <='" + k_d + "') as valuesmeters left join (SELECT  \"BalanceGroups_Meters\".\"BalanceGroupId\", \"BalanceGroups_Meters\".\"MeterId\", \"BalanceGroups\".\"Name\" FROM \"BalanceGroups_Meters\", \"BalanceGroups\" WHERE \"BalanceGroups\".\"BalanceGroupId\"=\"BalanceGroups_Meters\".\"BalanceGroupId\") as balancegroup on balancegroup.\"MeterId\"=valuesmeters.\"MeterId\" ", connection);
                //MessageBox.Show(" select  balancegroup.\"Name\", valuesmeters.\"SerialNumber\", valuesmeters.\"DT\", valuesmeters.\"Val\" from (SELECT  \"Meters\".\"SerialNumber\", \"Meters\".\"MeterId\",  \"Values\".\"DT\", \"Values\".\"Val\" from  \"Values\", \"Meters\" where \"Values\".\"MeterId\" IN (select \"MeterId\" from \"BalanceGroups_Meters\" where \"BalanceGroupId\" IN (select \"BalanceGroupId\" from \"BalanceGroups\" where \"Name\" IN (" + tii(clb) + "))) and \"Values\".\"MeterId\"=\"Meters\".\"MeterId\" and \"Values\".\"TariffId\" = 1 and \"Meters\".\"SerialNumber\" IN (select \"SerialNumber\" from \"Meters\" where \"ComObjectId\" IN (select \"ComObjectId\" from \"ComObjects\" where \"Name\" = '" + cb_rp.SelectedItem.ToString() + "')) and \"Values\".\"DT\" >= '" + n_d + "' and \"Values\".\"DT\" <= '" + k_d + "') as valuesmeters left join (SELECT  \"BalanceGroups_Meters\".\"BalanceGroupId\", \"BalanceGroups_Meters\".\"MeterId\", \"BalanceGroups\".\"Name\" FROM \"BalanceGroups_Meters\", \"BalanceGroups\" WHERE \"BalanceGroups\".\"BalanceGroupId\"=\"BalanceGroups_Meters\".\"BalanceGroupId\") as balancegroup on balancegroup.\"MeterId\"=valuesmeters.\"MeterId\" ");
                rtb.Text = rtb.Text + "select  balancegroup.\"Name\", valuesmeters.\"SerialNumber\", valuesmeters.\"DT\", valuesmeters.\"Val\" from (SELECT  \"Meters\".\"SerialNumber\", \"Meters\".\"MeterId\",  \"ValueProfiles\".\"DT\", \"ValueProfiles\".\"Val\" from  \"ValueProfiles\", \"Meters\" where \"ValueProfiles\".\"MeterId\" IN (select \"MeterId\" from \"BalanceGroups_Meters\" where \"BalanceGroupId\" IN (select \"BalanceGroupId\" from \"BalanceGroups\" where \"Name\" IN (" + tii(clb) + "))) and \"ValueProfiles\".\"MeterId\"=\"Meters\".\"MeterId\" and \"ValueProfiles\".\"StateId\" = 1 and \"Meters\".\"SerialNumber\" IN (select \"SerialNumber\" from \"Meters\" where \"ComObjectId\" IN (select \"ComObjectId\" from \"ComObjects\" where \"Name\" = '" + cb_rp.SelectedItem.ToString() + "')) and \"ValueProfiles\".\"DT\" >= '" + n_d + "' and \"ValueProfiles\".\"DT\" <='" + k_d + "') as valuesmeters left join (SELECT  \"BalanceGroups_Meters\".\"BalanceGroupId\", \"BalanceGroups_Meters\".\"MeterId\", \"BalanceGroups\".\"Name\" FROM \"BalanceGroups_Meters\", \"BalanceGroups\" WHERE \"BalanceGroups\".\"BalanceGroupId\"=\"BalanceGroups_Meters\".\"BalanceGroupId\") as balancegroup on balancegroup.\"MeterId\"=valuesmeters.\"MeterId\" ";
                NpgsqlDataAdapter dap = new NpgsqlDataAdapter(com);
                dap.Fill(table);
                connection.Close();

                foreach (DataRow row in table.Rows)
                {
                    energomera m = new energomera();
                    m.name_rp = cb_rp.SelectedItem.ToString();
                    m.name_sch = row.ItemArray[1].ToString();
                    m.name_f = row.ItemArray[0].ToString();
                    m.t_time = row.ItemArray[2].ToString();
                    m.v_value = Math.Round(Convert.ToDecimal(row.ItemArray[3]), 4).ToString();
                    vivod_energomera.Add(m);
                }

                dgv.DataSource = vivod_energomera;

                for (int i = 0; i < dgv.Rows.Count; i++)
                {
                    for (int j = 0; j < dgv.Columns.Count; j++)
                    {
                        for (int k = 0; k < energo_ktt.Count; k++)
                        {
                            if (dgv.Rows[i].Cells[1].Value.ToString().Equals(energo_ktt[k].fider.ToString())) { dgv.Rows[i].Cells[3].Value = energo_ktt[k].val; }
                        }
                    }
                }

                dgv.Columns[0].HeaderText = "Объект";
                dgv.Columns[1].HeaderText = "Присоединение";
                dgv.Columns[2].HeaderText = "Счетчик";
                dgv.Columns[3].HeaderText = "Кт";
                dgv.Columns[4].HeaderText = "Дата и время";
                dgv.Columns[5].HeaderText = "Мощность(кВт)";
            }

            NpgsqlConnection connection123 = new NpgsqlConnection(energomera_comobjects);
            connection123.Open();

            if (chb_energomera.Checked == true)
            {



                string rp = cb_rp.SelectedItem.ToString();
                List<string> list_f1 = new List<string>(); //лист фидеров
                NpgsqlConnection connection1 = new NpgsqlConnection(monitor_options);
                connection1.Open();
                NpgsqlCommand com1 = new NpgsqlCommand("select \"name_tii\", \"ktt\"  from \"balans\" where \"name_rp\" = '" + rp + "' ", connection1);
                rtb.Text = rtb.Text + "select \"name_tii\", \"ktt\"  from \"balans\" where \"name_rp\" = '" + rp + "' \n \n";
                NpgsqlDataReader datar1 = com1.ExecuteReader();
                fider.Clear();
                while (datar1.Read())
                {
                    class5 m = new class5();
                    m.fider = datar1.GetValue(0).ToString();
                    m.val = datar1.GetValue(1).ToString();
                    fider.Add(m);

                }
                connection1.Close();

                vivod_energomera.Clear();
                list_sutki_energo.Clear();

                byte dayCount = Convert.ToByte(x2 - x1);
                //MessageBox.Show(dayCount.ToString());
                for (byte i = 0; i < dayCount + 1; i++)
                {
                    for (int j = 0; j < list1_energo.Count; j++)
                    {
                        NpgsqlCommand com10 = new NpgsqlCommand("SELECT sum(\"ValueProfiles\".\"Val\") from  \"ValueProfiles\", \"Meters\" where \"ValueProfiles\".\"MeterId\" = " + list1_energo[j].meters_id + " and \"ValueProfiles\".\"MeterId\"=\"Meters\".\"MeterId\" and \"ValueProfiles\".\"StateId\" = 1 and \"ValueProfiles\".\"DT\" >= '" + DateTime.FromOADate(x1 + i) + "' and \"ValueProfiles\".\"DT\" < '" + DateTime.FromOADate(x1 + i + 1) + "'", connection123);
                        //rtb.Text = rtb.Text + "SELECT max(\"Values\".\"Val\") - min(\"Values\".\"Val\") from  \"Values\", \"Meters\" where \"Values\".\"MeterId\" = " + list1_energo[j].meters_id + " and \"Values\".\"MeterId\"=\"Meters\".\"MeterId\" and \"Values\".\"TariffId\" = 1 and \"Values\".\"DT\" >= '" + DateTime.FromOADate(x1 + i) + "' and \"Values\".\"DT\" < '" + DateTime.FromOADate(x1 + i + 1) + "'";
                        rtb.Text = rtb.Text + "SELECT sum(\"ValueProfiles\".\"Val\") from  \"ValueProfiles\", \"Meters\" where \"ValueProfiles\".\"MeterId\" = " + list1_energo[j].meters_id + " and \"ValueProfiles\".\"MeterId\"=\"Meters\".\"MeterId\" and \"ValueProfiles\".\"StateId\" = 1 and \"ValueProfiles\".\"DT\" >= '" + DateTime.FromOADate(x1 + i) + "' and \"ValueProfiles\".\"DT\" < '" + DateTime.FromOADate(x1 + i + 1) + "'";
                        com10.ExecuteNonQuery();
                        Decimal s1 = Convert.ToDecimal(com10.ExecuteScalar());
                        class_sutki_energo m = new class_sutki_energo();
                        m.id = list1_energo[j].meters_id;
                        m.name = list1_energo[j].name_f.ToString();
                        m.value = s1;
                        m.data = i;
                        list_sutki_energo.Add(m);
                    }
                }


                for (int i = 0; i < list_sutki_energo.Count; i++)
                {
                    for (int j = 0; j < list1_energo.Count; j++)
                    {
                        if (list_sutki_energo[i].id == list1_energo[j].meters_id)
                        {

                            energomera m = new energomera();
                            m.name_rp = list1_energo[j].name;
                            m.name_f = list1_energo[j].name_f;
                            m.name_sch = list1_energo[j].sn;
                            m.t_time = DateTime.FromOADate((x1 + list_sutki_energo[i].data)).ToString();
                            m.v_value = Math.Round(list_sutki_energo[i].value, 3).ToString();
                            for (int k = 0; k < fider.Count; k++)
                            {

                                if (list1_energo[j].name_f.Equals(fider[k].fider))
                                { m.ktt = fider[k].val; }


                            }

                            vivod_energomera.Add(m);
                        }
                    }
                }
                energomera m1 = new energomera();
                m1.name_rp = "Итог";
                m1.t_time = "За период с " + d1 + " до " + d2;
                m1.name_f = "";
                vivod_energomera.Add(m1);

                dgv.DataSource = vivod_energomera;

                for (int i = 0; i < dgv.Rows.Count; i++)
                {

                    dgv.Rows[i].Cells[5].Value = ((Convert.ToDecimal(dgv.Rows[i].Cells[5].Value) * Convert.ToDecimal(dgv.Rows[i].Cells[3].Value)) / 2).ToString();

                }



                List<string> s_list = new List<string>();
                List<string[]> sum_list = new List<string[]>();
                s_list.Clear();
                sum_list.Clear();
                string stroka = "";
                string strochka = "";
                string strochka1 = "";
                for (int i = 0; i < vivod_energomera.Count - 1; i++)
                {

                    if (!s_list.Contains(vivod_energomera[i].name_f)) { s_list.Add(vivod_energomera[i].name_f); }
                }

                for (int k = 0; k < s_list.Count; k++)
                {
                    string[] str = { s_list[k], "0" };
                    sum_list.Add(str);
                }

                for (int k = 0; k < sum_list.Count; k++)
                {
                    for (int i = 0; i < vivod_energomera.Count; i++)
                    {
                        if (sum_list[k][0].Equals(vivod_energomera[i].name_f))
                        {
                            sum_list[k][1] = Math.Round((Convert.ToDecimal(sum_list[k][1]) + Convert.ToDecimal(vivod_energomera[i].v_value)), 3).ToString();
                        }

                    }
                }


                for (int j = 0; j < sum_list.Count; j++)
                {
                    stroka = stroka + sum_list[j][0].ToString() + " = " + sum_list[j][1].ToString() + "\n";
                    strochka1 = strochka1 + sum_list[j][0].ToString() + "\n";
                    strochka = strochka + sum_list[j][1].ToString() + "\n";

                }
                for (int i = 0; i < dgv.Rows.Count; i++)
                {
                    //dataGridView3.Rows.Add();

                    //dgv.Rows[dgv.Rows.Count - 1].Cells[5].Value = "ИТОГО";
                    dgv.Rows[dgv.Rows.Count - 1].Cells[5].Value = "\n" + stroka;
                }



                dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells); //автоматический размер колонок
                dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                dgv.Columns[0].HeaderText = "Объект";
                dgv.Columns[1].HeaderText = "Присоединение"; //присваиваем текст к заголовку колонок
                dgv.Columns[2].HeaderText = "Счетчик №";
                dgv.Columns[3].HeaderText = "Кт";
                dgv.Columns[4].HeaderText = "Дата и время";
                dgv.Columns[5].HeaderText = "Энергия(кВт*ч)";

            }
            return dgv;

        } //результат данных по энергомере
        public Chart graphic(Chart ch, DataGridView dgv, Int32 number_cell_y, Int32 number_cell_x, Int32 number_cell_y1, Int32 number_cell_x1, string format, string y, string x, CheckedListBox clb, DateTimePicker dtp_n, DateTimePicker dtp_k, ComboBox cb_base)
        {
            DateTime d1 = new DateTime(dtp_n.Value.Year, dtp_n.Value.Month, dtp_n.Value.Day); //дата начала периода
            double nd = d1.ToOADate();    //перевод нормальной даты в абсолютную
            DateTime d2 = new DateTime(dtp_k.Value.Year, dtp_k.Value.Month, dtp_k.Value.Day); //дата конца периода
            double kd = d2.ToOADate();    //перевод нормальной даты в абсолютную

            ch.Visible = true;
            ch.BringToFront();
            double maxval = -1;
            if (cb_base.SelectedItem.ToString() == "Компас РП")
            {




                for (int k = 0; k < dgv.Rows.Count - 1; k++)
                {
                    if (maxval < Convert.ToDouble(dgv.Rows[k].Cells[number_cell_y].Value))
                    { maxval = Convert.ToDouble(dgv.Rows[k].Cells[number_cell_y].Value); }

                }
                //label1.Text = maxval.ToString();
                //sum1(sender, e, form.dataGridView3, label4);
                int ss = 0;
                ch.Series[0].Points.Clear();
                while (ss < dgv.Rows.Count - 1)
                {
                    var skip = dgv.Rows[ss].Cells[number_cell_y].Value;
                    var date = dgv.Rows[ss].Cells[number_cell_x].Value;
                    //string s = Convert.ToString(date);
                    //string s1 = s.Substring(s.IndexOf(" "),8);
                    //chart1.ChartAreas[0].AxisY.MajorGrid.Interval = 50;
                    ch.ChartAreas[0].AxisX.LabelStyle.Format = format;
                    ch.ChartAreas[0].AxisX.IntervalType = System.Windows.Forms.DataVisualization.Charting.DateTimeIntervalType.Days;
                    ch.ChartAreas[0].AxisX.Interval = 1;
                    ch.Titles[0].Text = "График энергии   " + tii(clb) + "   за период   " + d1 + "   -   " + d2;
                    foreach (Legend objLegend in ch.Legends) //убрали легенду из графика
                    {
                        objLegend.Enabled = false;
                    }

                    ch.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                    ch.Series[0].Points.AddXY(Convert.ToDateTime(date), Convert.ToDouble(skip));
                    ch.Series[0].ToolTip = "X=#VALX, Y=#VALY";
                    ch.ChartAreas[0].AxisY.Title = y;
                    ch.ChartAreas[0].AxisX.Title = x;
                    ss++;
                }
            }

            if (cb_base.SelectedItem.ToString() == "Энергомера")
            {
                for (int k = 0; k < dgv.Rows.Count - 1; k++)
                {
                    if (maxval < Convert.ToDouble(dgv.Rows[k].Cells[number_cell_y1].Value))
                    { maxval = Convert.ToDouble(dgv.Rows[k].Cells[number_cell_y1].Value); }

                }
                //label1.Text = maxval.ToString();
                //sum1(sender, e, form.dataGridView3, label4);
                int ss = 0;
                ch.Series[0].Points.Clear();
                while (ss < dgv.Rows.Count - 1)
                {
                    var skip = dgv.Rows[ss].Cells[number_cell_y1].Value;
                    var date = dgv.Rows[ss].Cells[number_cell_x1].Value;
                    //string s = Convert.ToString(date);
                    //string s1 = s.Substring(s.IndexOf(" "),8);
                    //chart1.ChartAreas[0].AxisY.MajorGrid.Interval = 50;
                    ch.ChartAreas[0].AxisX.LabelStyle.Format = format;
                    ch.ChartAreas[0].AxisX.IntervalType = System.Windows.Forms.DataVisualization.Charting.DateTimeIntervalType.Days;
                    ch.ChartAreas[0].AxisX.Interval = 1;
                    ch.Titles[0].Text = "График энергии   " + tii(clb) + "   за период   " + d1 + "   -   " + d2;
                    foreach (Legend objLegend in ch.Legends) //убрали легенду из графика
                    {
                        objLegend.Enabled = false;
                    }

                    ch.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                    ch.Series[0].Points.AddXY(Convert.ToDateTime(date), Convert.ToDouble(skip));
                    ch.Series[0].ToolTip = "X=#VALX, Y=#VALY";
                    ch.ChartAreas[0].AxisY.Title = y;
                    ch.ChartAreas[0].AxisX.Title = x;
                    ss++;
                }
            }

            return ch;
        } //График
        public CheckBox videl_pok(CheckBox cb, CheckedListBox clb)
        {
            if (cb.Checked == true)
            {

                for (int i = 0; i < clb.Items.Count; i++)
                {

                    if (clb.Items[i].ToString().Substring(0, 3) == "пок" || clb.Items[i].ToString().Substring(0, 3) == "Пок")
                    { clb.SetItemChecked(i, true); }

                }

            }
            if (cb.Checked == false)
            {

                for (int i = 0; i < clb.Items.Count; i++)
                {

                    if (clb.Items[i].ToString().Substring(0, 3) == "пок" || clb.Items[i].ToString().Substring(0, 3) == "Пок")
                    { clb.SetItemChecked(i, false); }

                }
            }
            return cb;
        } //Выделение Пок
        public CheckBox videl_tp(CheckBox cb, CheckedListBox clb)
        {
            if (cb.Checked == true)
            {

                for (int i = 0; i < clb.Items.Count; i++)
                {

                    if (clb.Items[i].ToString().Substring(0, 2) == "ТП" || clb.Items[i].ToString().Substring(0, 2) == "тп")
                    { clb.SetItemChecked(i, true); }
                }

            }

            if (cb.Checked == false)
            {

                for (int i = 0; i < clb.Items.Count; i++)
                {

                    if (clb.Items[i].ToString().Substring(0, 2) == "ТП" || clb.Items[i].ToString().Substring(0, 2) == "тп")
                    { clb.SetItemChecked(i, false); }
                }

            }
            return cb;

        } //Выделение ТП
        public CheckBox videl_p(CheckBox cb, CheckedListBox clb)
        {
            if (cb.Checked == true)
            {

                for (int i = 0; i < clb.Items.Count; i++)
                {

                    if (clb.Items[i].ToString().Substring(0, 1) == "P" || clb.Items[i].ToString().Substring(0, 1) == "Р")
                    { clb.SetItemChecked(i, true); }
                }

            }

            if (cb.Checked == false)
            {

                for (int i = 0; i < clb.Items.Count; i++)
                {

                    if (clb.Items[i].ToString().Substring(0, 1) == "P" || clb.Items[i].ToString().Substring(0, 1) == "Р")
                    { clb.SetItemChecked(i, false); }
                }

            }
            return cb;

        } //Выделение P
        public DataGridView nebalans_vis(DataGridView dgv, DateTimePicker dtp_n, DateTimePicker dtp_k, ComboBox cb_rp, ContextMenuStrip cms, Chart ch, ComboBox cb_base, RichTextBox rtb)
        {
            dgv.DataSource = null;
            list_nebalans.Clear();
            list_f.Clear();
            list_v.Clear();
            list_vn.Clear();
            list_fn.Clear();
            cms.Items[4].Enabled = true;
            cms.Items[6].Enabled = false;
            cms.Items[7].Enabled = false;


            ktt(cb_rp);

            string rp = cb_rp.SelectedItem.ToString();
            DateTime d1 = new DateTime(dtp_n.Value.Year, dtp_n.Value.Month, dtp_n.Value.Day); //дата начала периода
            double nd = d1.ToOADate();    //перевод нормальной даты в абсолютную
            DateTime d2 = new DateTime(dtp_k.Value.Year, dtp_k.Value.Month, dtp_k.Value.Day); //дата конца периода
            double kd = d2.ToOADate();    //перевод нормальной даты в абсолютную

            long base_tick = new DateTime(1601, 1, 1).Ticks;
            long n_tick = d1.Ticks;
            long n_tick_mysql = n_tick - base_tick;
            long k_tick = d2.Ticks;
            long k_tick_mysql = k_tick - base_tick;

            //if (kd - nd == 0) { MessageBox.Show("Выбрать дату"); return; }
            long dayCount = k_tick_mysql - n_tick_mysql;
            long oneday = 864000000000;
            //byte dayCount = Convert.ToByte(kd - nd);
            //MessageBox.Show(DateTime.FromOADate(nd+dayCount).ToString());

            string mes_n = dtp_n.Value.Month.ToString();
            string mes_k = dtp_k.Value.Month.ToString();


            switch (mes_n)
            {
                case "1": mes1 = "Январь"; break;
                case "2": mes1 = "Февраль"; break;
                case "3": mes1 = "Март"; break;
                case "4": mes1 = "Апрель"; break;
                case "5": mes1 = "Май"; break;
                case "6": mes1 = "Июнь"; break;
                case "7": mes1 = "Июль"; break;
                case "8": mes1 = "Август"; break;
                case "9": mes1 = "Сентябрь"; break;
                case "10": mes1 = "Октябрь"; break;
                case "11": mes1 = "Ноябрь"; break;
                case "12": mes1 = "Декабрь"; break;

            }

            switch (mes_k)
            {
                case "1": mes2 = "Январь"; break;
                case "2": mes2 = "Февраль"; break;
                case "3": mes2 = "Март"; break;
                case "4": mes2 = "Апрель"; break;
                case "5": mes2 = "Май"; break;
                case "6": mes2 = "Июнь"; break;
                case "7": mes2 = "Июль"; break;
                case "8": mes2 = "Август"; break;
                case "9": mes2 = "Сентябрь"; break;
                case "10": mes2 = "Октябрь"; break;
                case "11": mes2 = "Ноябрь"; break;
                case "12": mes2 = "Декабрь"; break;

            }

            if (cb_base.SelectedItem.ToString() == "Компас РП")
            {

                NpgsqlConnection connection = new NpgsqlConnection(monitor_options);
                NpgsqlConnection connection1 = new NpgsqlConnection(monitor);
                connection1.Open();
                connection.Open();
                NpgsqlCommand com = new NpgsqlCommand("select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and \"vvod\" = TRUE", connection);
                rtb.Text = rtb.Text + "Небаланс Выс \n \n" + "select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and \"vvod\" = TRUE \n \n";
                NpgsqlDataReader datar = com.ExecuteReader();

                while (datar.Read())
                {

                    list_v.Add(datar.GetValue(0).ToString());
                    list_vn.Add(datar.GetValue(0).ToString());

                }

                com = new NpgsqlCommand("select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and \"fider\" = TRUE", connection);
                rtb.Text = rtb.Text + "select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and \"fider\" = TRUE \n \n";
                datar = com.ExecuteReader();

                while (datar.Read())
                {
                    list_f.Add(datar.GetValue(0).ToString());
                    list_fn.Add(datar.GetValue(0).ToString());

                }

                string g_id = string.Empty;
                NpgsqlCommand com2 = new NpgsqlCommand("select t1.\"Groups_ID\" from \"GROUPS\" t1, \"GROUPS\" t2 where t1.\"Tag\" = 'ТИИ' and t1.\"Parent\"=t2.\"Groups_ID\" and t2.\"Name\" = '" + cb_rp.SelectedItem.ToString() + "'", connection1);
                //MessageBox.Show("select t1.\"Groups_ID\" from \"GROUPS\" t1, \"GROUPS\" t2 where t1.\"Tag\" = 'ТИИ' and t1.\"Parent\"=t2.\"Groups_ID\" and t2.\"Name\" = '" + form5.comboBox1.SelectedItem.ToString() + "'");
                rtb.Text = rtb.Text + "select t1.\"Groups_ID\" from \"GROUPS\" t1, \"GROUPS\" t2 where t1.\"Tag\" = 'ТИИ' and t1.\"Parent\"=t2.\"Groups_ID\" and t2.\"Name\" = '" + cb_rp.SelectedItem.ToString() + "' \n \n";
                g_id = Convert.ToString(com2.ExecuteScalar());


                for (int i = 0; i < list_v.Count; i++)
                {
                    NpgsqlCommand com3 = new NpgsqlCommand("select \"Values_ID\" from \"TB_VALUES\" where \"Groups_ID\" =" + g_id + " and \"Name\" = '" + list_v[i] + "' ", connection1);
                    //MessageBox.Show("select \"Values_ID\" from \"TB_VALUES\" where \"Groups_ID\" =" + g_id + " and \"Name\" = '" + list_v[i] + "' ");
                    rtb.Text = rtb.Text + "select \"Values_ID\" from \"TB_VALUES\" where \"Groups_ID\" =" + g_id + " and \"Name\" = '" + list_v[i] + "' \n \n";
                    list_v[i] = Convert.ToString(com3.ExecuteScalar());
                }


                for (int i = 0; i < list_f.Count; i++)
                {
                    NpgsqlCommand com4 = new NpgsqlCommand("select \"Values_ID\" from \"TB_VALUES\" where \"Groups_ID\" =" + g_id + " and \"Name\" = '" + list_f[i] + "' ", connection1);
                    rtb.Text = rtb.Text + "select \"Values_ID\" from \"TB_VALUES\" where \"Groups_ID\" =" + g_id + " and \"Name\" = '" + list_f[i] + "' \n \n";
                    list_f[i] = Convert.ToString(com4.ExecuteScalar());
                }

                connection1.Close();
                connection.Close();
            }

            if (cb_base.SelectedItem.ToString() == "Энергомера")
            {
                NpgsqlConnection connection = new NpgsqlConnection(monitor_options);
                NpgsqlConnection connection1 = new NpgsqlConnection(energomera_comobjects);
                connection.Open();
                connection1.Open();
                NpgsqlCommand com = new NpgsqlCommand("select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and \"vvod\" = TRUE", connection);
                rtb.Text = rtb.Text + "Небаланс период \n \n" + "select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and \"vvod\" = TRUE \n \n";
                NpgsqlDataReader datar = com.ExecuteReader();

                while (datar.Read())
                {

                    list_v.Add(datar.GetValue(0).ToString());
                    list_vn.Add(datar.GetValue(0).ToString());

                }

                com = new NpgsqlCommand("select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and \"fider\" = TRUE ", connection);
                rtb.Text = rtb.Text + "select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and \"fider\" = TRUE \n \n";
                datar = com.ExecuteReader();

                while (datar.Read())
                {
                    list_f.Add(datar.GetValue(0).ToString());
                    list_fn.Add(datar.GetValue(0).ToString());

                }

                for (int i = 0; i < list_v.Count; i++)
                {
                    NpgsqlCommand com3 = new NpgsqlCommand("select \"MeterId\" from \"BalanceGroups_Meters\" where \"BalanceGroupId\" IN (select \"BalanceGroupId\" from \"BalanceGroups\" where \"Name\" IN ('" + list_v[i] + "'))", connection1);
                    //MessageBox.Show("select \"Values_ID\" from \"TB_VALUES\" where \"Groups_ID\" =" + g_id + " and \"Name\" = '" + list_v[i] + "' ");
                    rtb.Text = rtb.Text + "select \"MeterId\" from \"BalanceGroups_Meters\" where \"BalanceGroupId\" IN (select \"BalanceGroupId\" from \"BalanceGroups\" where \"Name\" IN ('" + list_v[i] + "')) \n \n";
                    list_v[i] = Convert.ToString(com3.ExecuteScalar());
                }


                for (int i = 0; i < list_f.Count; i++)
                {
                    NpgsqlCommand com4 = new NpgsqlCommand("select \"MeterId\" from \"BalanceGroups_Meters\" where \"BalanceGroupId\" IN (select \"BalanceGroupId\" from \"BalanceGroups\" where \"Name\" IN ('" + list_f[i] + "'))", connection1);
                    rtb.Text = rtb.Text + "select \"MeterId\" from \"BalanceGroups_Meters\" where \"BalanceGroupId\" IN (select \"BalanceGroupId\" from \"BalanceGroups\" where \"Name\" IN ('" + list_f[i] + "')) \n \n";
                    list_f[i] = Convert.ToString(com4.ExecuteScalar());
                }


                connection.Close();
                connection1.Close();
            }


            if (cms.Items[2].Text == "Небаланс РУ-6-10 кВ")
            {
                b.Clear();
                b1.Clear();
                dat.Clear();

                //for (byte i = 0; i <= dayCount; i++)
                    for (long i = 0; i <= dayCount; i = i + oneday)
                    {
                        //MessageBox.Show(i.ToString() + "--" + dayCount);
                    long x1 = Convert.ToInt64(i + base_tick);
                    DateTime dt1 = new DateTime(x1);
                    //dat.Add(new DateTime(dtp_n.Value.Year, dtp_n.Value.Month, dtp_n.Value.Day).AddDays(i));
                    dat.Add(new DateTime(dtp_n.Value.Year, dtp_n.Value.Month, dtp_n.Value.Day).AddDays(dt1.Day-1));
                    decimal sum_v1 = 0;
                    decimal sum_f1 = 0;

                    for (int k = 0; k < list_v.Count; k++)
                    {
                        sum_v1 = sum_v1 + summ(list_v[k], i, list_vn[k], dtp_n, cb_base, rtb);
                    }

                    for (int k = 0; k < list_f.Count; k++)
                    {
                        sum_f1 = sum_f1 + summ(list_f[k], i, list_fn[k], dtp_n, cb_base, rtb);
                    }
                    //MessageBox.Show(sum_v1.ToString() + " - " + sum_f1.ToString());
                    b.Add(Math.Round((sum_v1 - sum_f1), 3));
                    try
                    {
                        b1.Add(Math.Round(((sum_v1 - sum_f1) / sum_v1) * 100, 3));
                    }
                    catch { b1.Add(0); }
                    //catch { MessageBox.Show("Сумма вводов равна нулю. На ноль делить нельзя. \n Проверьте показания."); } 


                    class_nebalans m = new class_nebalans();
                    m.data = dat[dt1.Day-1].ToString();
                    m.value = b[dt1.Day-1].ToString();
                    m.proc = b1[dt1.Day-1].ToString();
                    list_nebalans.Add(m);



                }

            }
            dgv.DataSource = list_nebalans;
            dgv.Columns[0].HeaderText = "Дата";
            dgv.Columns[1].HeaderText = "кВт*ч";
            dgv.Columns[2].HeaderText = "%";
            dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
            int ss = 0;
            ch.Series[0].Points.Clear();
            while (ss < b.Count)
            {

                var proc = b1[ss];
                var skip = b[ss];
                var date = dat[ss];
                //chart1.ChartAreas[0].AxisY.MajorGrid.Interval = 50;
                ch.ChartAreas[0].AxisX.LabelStyle.Format = "dd";
                ch.ChartAreas[0].AxisX.IntervalType = System.Windows.Forms.DataVisualization.Charting.DateTimeIntervalType.Days;
                ch.ChartAreas[0].AxisX.Interval = 1;
                ch.Titles[0].Text = "График небаланса   " + cb_rp.SelectedItem.ToString() + "   за период   " + d1 + "   -   " + d2 + ru_vis;
                foreach (Legend objLegend in ch.Legends) //убрали легенду из графика
                {
                    objLegend.Enabled = false;
                }
                ch.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                ch.Series[0].Points.AddXY(Convert.ToDateTime(date), Convert.ToDouble(skip));
                ch.Series[0].ToolTip = "X=#VALX, Y=#VALY";
                ss++;
            }


            return dgv;
        } //Небаланс по высокой + график
        public DataGridView nebalans_niz(DataGridView dgv, DateTimePicker dtp_n, DateTimePicker dtp_k, ComboBox cb_rp, ContextMenuStrip cms, Chart ch, ComboBox cb_base, RichTextBox rtb)
        {

            dgv.DataSource = null;
            list_nebalans.Clear();

            list_v_niz.Clear();
            list_f_niz.Clear();
            list_vn_niz.Clear();
            list_fn_niz.Clear();
            cms.Items[4].Enabled = true;
            cms.Items[6].Enabled = false;
            cms.Items[7].Enabled = false;

            ktt(cb_rp);

            string rp = cb_rp.SelectedItem.ToString();
            DateTime d1 = new DateTime(dtp_n.Value.Year, dtp_n.Value.Month, dtp_n.Value.Day); //дата начала периода
            double nd = d1.ToOADate();    //перевод нормальной даты в абсолютную
            DateTime d2 = new DateTime(dtp_k.Value.Year, dtp_k.Value.Month, dtp_k.Value.Day); //дата конца периода
            double kd = d2.ToOADate();    //перевод нормальной даты в абсолютную

            long base_tick = new DateTime(1601, 1, 1).Ticks;
            long n_tick = d1.Ticks;
            long n_tick_mysql = n_tick - base_tick;
            long k_tick = d2.Ticks;
            long k_tick_mysql = k_tick - base_tick;

            //if (kd - nd == 0) { MessageBox.Show("Выбрать дату"); return; }

            //byte dayCount = Convert.ToByte(kd - nd);

            long dayCount = k_tick_mysql - n_tick_mysql;
            long oneday = 864000000000;

            string mes_n = dtp_n.Value.Month.ToString();
            string mes_k = dtp_k.Value.Month.ToString();


            switch (mes_n)
            {
                case "1": mes1 = "Январь"; break;
                case "2": mes1 = "Февраль"; break;
                case "3": mes1 = "Март"; break;
                case "4": mes1 = "Апрель"; break;
                case "5": mes1 = "Май"; break;
                case "6": mes1 = "Июнь"; break;
                case "7": mes1 = "Июль"; break;
                case "8": mes1 = "Август"; break;
                case "9": mes1 = "Сентябрь"; break;
                case "10": mes1 = "Октябрь"; break;
                case "11": mes1 = "Ноябрь"; break;
                case "12": mes1 = "Декабрь"; break;

            }

            switch (mes_k)
            {
                case "1": mes2 = "Январь"; break;
                case "2": mes2 = "Февраль"; break;
                case "3": mes2 = "Март"; break;
                case "4": mes2 = "Апрель"; break;
                case "5": mes2 = "Май"; break;
                case "6": mes2 = "Июнь"; break;
                case "7": mes2 = "Июль"; break;
                case "8": mes2 = "Август"; break;
                case "9": mes2 = "Сентябрь"; break;
                case "10": mes2 = "Октябрь"; break;
                case "11": mes2 = "Ноябрь"; break;
                case "12": mes2 = "Декабрь"; break;

            }

            if (cb_base.SelectedItem.ToString() == "Компас РП")
            {

                NpgsqlConnection connection = new NpgsqlConnection(monitor_options);
                NpgsqlConnection connection1 = new NpgsqlConnection(monitor);
                connection1.Open();
                connection.Open();
                NpgsqlCommand com = new NpgsqlCommand("select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and \"vvod_niz\" = TRUE", connection);
                NpgsqlDataReader datar = com.ExecuteReader();

                while (datar.Read())
                {
                    list_v_niz.Add(datar.GetValue(0).ToString());
                    list_vn_niz.Add(datar.GetValue(0).ToString());

                }

                com = new NpgsqlCommand("select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and \"fider_niz\" = TRUE", connection);
                rtb.Text = rtb.Text + "Небаланс Низ \n \n" + "select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and \"fider_niz\" = TRUE \n \n";
                datar = com.ExecuteReader();

                while (datar.Read())
                {
                    list_f_niz.Add(datar.GetValue(0).ToString());
                    list_fn_niz.Add(datar.GetValue(0).ToString());

                }


                string g_id = string.Empty;
                NpgsqlCommand com2 = new NpgsqlCommand("select t1.\"Groups_ID\" from \"GROUPS\" t1, \"GROUPS\" t2 where t1.\"Tag\" = 'ТИИ' and t1.\"Parent\"=t2.\"Groups_ID\" and t2.\"Name\" = '" + cb_rp.SelectedItem.ToString() + "'", connection1);
                //MessageBox.Show("select t1.\"Groups_ID\" from \"GROUPS\" t1, \"GROUPS\" t2 where t1.\"Tag\" = 'ТИИ' and t1.\"Parent\"=t2.\"Groups_ID\" and t2.\"Name\" = '" + form5.comboBox1.SelectedItem.ToString() + "'");
                rtb.Text = rtb.Text + "select t1.\"Groups_ID\" from \"GROUPS\" t1, \"GROUPS\" t2 where t1.\"Tag\" = 'ТИИ' and t1.\"Parent\"=t2.\"Groups_ID\" and t2.\"Name\" = '" + cb_rp.SelectedItem.ToString() + "' \n \n";
                g_id = Convert.ToString(com2.ExecuteScalar());

                for (int i = 0; i < list_v_niz.Count; i++)
                {
                    NpgsqlCommand com5 = new NpgsqlCommand("select \"Values_ID\" from \"TB_VALUES\" where \"Groups_ID\" =" + g_id + " and \"Name\" = '" + list_v_niz[i] + "' ", connection1);
                    //MessageBox.Show("select \"Values_ID\" from \"TB_VALUES\" where \"Groups_ID\" =" + g_id + " and \"Name\" = '" + list_v_niz[i] + "' ");
                    rtb.Text = rtb.Text + "select \"Values_ID\" from \"TB_VALUES\" where \"Groups_ID\" =" + g_id + " and \"Name\" = '" + list_v_niz[i] + "' \n \n";
                    list_v_niz[i] = Convert.ToString(com5.ExecuteScalar());
                }

                for (int i = 0; i < list_f_niz.Count; i++)
                {
                    NpgsqlCommand com6 = new NpgsqlCommand("select \"Values_ID\" from \"TB_VALUES\" where \"Groups_ID\" =" + g_id + " and \"Name\" = '" + list_f_niz[i] + "' ", connection1);
                    //MessageBox.Show("select \"Values_ID\" from \"TB_VALUES\" where \"Groups_ID\" =" + g_id + " and \"Name\" = '" + list_f_niz[i] + "' ");
                    rtb.Text = rtb.Text + "select \"Values_ID\" from \"TB_VALUES\" where \"Groups_ID\" =" + g_id + " and \"Name\" = '" + list_f_niz[i] + "' \n \n";
                    list_f_niz[i] = Convert.ToString(com6.ExecuteScalar());
                }


                connection1.Close();
                connection.Close();
            }

            if (cb_base.SelectedItem.ToString() == "Энергомера")
            {
                NpgsqlConnection connection = new NpgsqlConnection(monitor_options);
                NpgsqlConnection connection1 = new NpgsqlConnection(energomera_comobjects);
                connection.Open();
                connection1.Open();
                NpgsqlCommand com = new NpgsqlCommand("select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and \"vvod_niz\" = TRUE", connection);
                rtb.Text = rtb.Text + "Небаланс период \n \n" + "select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and \"vvod_niz\" = TRUE \n \n";
                NpgsqlDataReader datar = com.ExecuteReader();

                while (datar.Read())
                {

                    list_v_niz.Add(datar.GetValue(0).ToString());
                    list_vn_niz.Add(datar.GetValue(0).ToString());

                }

                com = new NpgsqlCommand("select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and \"fider_niz\" = TRUE ", connection);
                rtb.Text = rtb.Text + "select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and \"fider_niz\" = TRUE \n \n";
                datar = com.ExecuteReader();

                while (datar.Read())
                {
                    list_f_niz.Add(datar.GetValue(0).ToString());
                    list_fn_niz.Add(datar.GetValue(0).ToString());

                }

                for (int i = 0; i < list_v_niz.Count; i++)
                {
                    NpgsqlCommand com3 = new NpgsqlCommand("select \"MeterId\" from \"BalanceGroups_Meters\" where \"BalanceGroupId\" IN (select \"BalanceGroupId\" from \"BalanceGroups\" where \"Name\" IN ('" + list_v_niz[i] + "'))", connection1);
                    //MessageBox.Show("select \"Values_ID\" from \"TB_VALUES\" where \"Groups_ID\" =" + g_id + " and \"Name\" = '" + list_v[i] + "' ");
                    rtb.Text = rtb.Text + "select \"MeterId\" from \"BalanceGroups_Meters\" where \"BalanceGroupId\" IN (select \"BalanceGroupId\" from \"BalanceGroups\" where \"Name\" IN ('" + list_v_niz[i] + "')) \n \n";
                    list_v_niz[i] = Convert.ToString(com3.ExecuteScalar());
                }


                for (int i = 0; i < list_f_niz.Count; i++)
                {
                    NpgsqlCommand com4 = new NpgsqlCommand("select \"MeterId\" from \"BalanceGroups_Meters\" where \"BalanceGroupId\" IN (select \"BalanceGroupId\" from \"BalanceGroups\" where \"Name\" IN ('" + list_f_niz[i] + "'))", connection1);
                    rtb.Text = rtb.Text + "select \"MeterId\" from \"BalanceGroups_Meters\" where \"BalanceGroupId\" IN (select \"BalanceGroupId\" from \"BalanceGroups\" where \"Name\" IN ('" + list_f_niz[i] + "')) \n \n";
                    list_f_niz[i] = Convert.ToString(com4.ExecuteScalar());
                }


                connection.Close();
                connection1.Close();
            }


            if (cms.Items[3].Text == "Небаланс РУ-0,4 кВ")
            {
                b.Clear();
                b1.Clear();
                dat.Clear();

               // for (byte i = 0; i <= dayCount; i++)
                for (long i = 0; i <= dayCount; i = i + oneday)
                {
                    long x1 = Convert.ToInt64(i + base_tick);
                    DateTime dt1 = new DateTime(x1);
                    //dat.Add(new DateTime(dtp_n.Value.Year, dtp_n.Value.Month, dtp_n.Value.Day).AddDays(i));
                    dat.Add(new DateTime(dtp_n.Value.Year, dtp_n.Value.Month, dtp_n.Value.Day).AddDays(dt1.Day - 1));

                    decimal sum_v1 = 0;
                    decimal sum_f1 = 0;

                    for (int k = 0; k < list_v_niz.Count; k++)
                    {
                        sum_v1 = sum_v1 + summ(list_v_niz[k], i, list_vn_niz[k], dtp_n, cb_base, rtb);

                    }

                    for (int k = 0; k < list_f_niz.Count; k++)
                    {
                        sum_f1 = sum_f1 + summ(list_f_niz[k], i, list_fn_niz[k], dtp_n, cb_base, rtb);
                    }
                    //MessageBox.Show(sum_v1.ToString() + " - " + sum_f1.ToString());

                    b.Add(Math.Round((sum_v1 - sum_f1), 3));
                    try
                    {
                        b1.Add(Math.Round(((sum_v1 - sum_f1) / sum_v1) * 100, 3));
                    }
                    catch { b1.Add(0); }
                    //catch { MessageBox.Show("Сумма вводов равна нулю. На ноль делить нельзя. \n Проверьте показания."); return; }  


                    class_nebalans m = new class_nebalans();
                    m.data = dat[dt1.Day - 1].ToString();
                    m.value = b[dt1.Day - 1].ToString();
                    m.proc = b1[dt1.Day - 1].ToString();
                    list_nebalans.Add(m);


                }

            }
            dgv.DataSource = list_nebalans;
            dgv.Columns[0].HeaderText = "Дата";
            dgv.Columns[1].HeaderText = "кВт*ч";
            dgv.Columns[2].HeaderText = "%";
            dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);

            int ss = 0;
            ch.Series[0].Points.Clear();
            while (ss < b.Count)
            {

                var proc = b1[ss];
                var skip = b[ss];
                var date = dat[ss];
                //chart1.ChartAreas[0].AxisY.MajorGrid.Interval = 50;
                ch.ChartAreas[0].AxisX.LabelStyle.Format = "dd";
                ch.ChartAreas[0].AxisX.IntervalType = System.Windows.Forms.DataVisualization.Charting.DateTimeIntervalType.Days;
                ch.ChartAreas[0].AxisX.Interval = 1;
                ch.Titles[0].Text = "График небаланса   " + cb_rp.SelectedItem.ToString() + "   за период   " + d1 + "   -   " + d2 + ru_niz;
                foreach (Legend objLegend in ch.Legends) //убрали легенду из графика
                {
                    objLegend.Enabled = false;
                }
                ch.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                ch.Series[0].Points.AddXY(Convert.ToDateTime(date), Convert.ToDouble(skip));
                ch.Series[0].ToolTip = "X=#VALX, Y=#VALY";
                ss++;
            }


            return dgv;
        } //Небаланс по низкой + график
        public DataGridView nebalans_vis_per(DataGridView dgv, DateTimePicker dtp_n, DateTimePicker dtp_k, ComboBox cb_rp, ContextMenuStrip cms, Chart ch, ComboBox cb_base, RichTextBox rtb)
        {
            dgv.DataSource = null;
            list_nebalans.Clear();
            list_f.Clear();
            list_v.Clear();
            list_vn.Clear();
            list_fn.Clear();
            //cms.Items[4].Enabled = true;
            cms.Items[6].Enabled = false;
            cms.Items[7].Enabled = false;


            ktt(cb_rp);

            string rp = cb_rp.SelectedItem.ToString();
            DateTime d1 = new DateTime(dtp_n.Value.Year, dtp_n.Value.Month, dtp_n.Value.Day); //дата начала периода
            double nd = d1.ToOADate();    //перевод нормальной даты в абсолютную
            DateTime d2 = new DateTime(dtp_k.Value.Year, dtp_k.Value.Month, dtp_k.Value.Day); //дата конца периода
            double kd = d2.ToOADate();    //перевод нормальной даты в абсолютную

            //if (kd - nd == 0) { MessageBox.Show("Выбрать дату"); return; }

            // byte dayCount = Convert.ToByte(kd - nd);

            //MessageBox.Show(d1.ToString() + "   " + nd.ToString() + "   " + d2.ToString() + "    " + kd.ToString() + "   " + dayCount.ToString());

            //string mes_n = dtp_n.Value.Month.ToString();
            //string mes_k = dtp_k.Value.Month.ToString();


            //switch (mes_n)
            //{
            //    case "1": mes1 = "Январь"; break;
            //    case "2": mes1 = "Февраль"; break;
            //    case "3": mes1 = "Март"; break;
            //    case "4": mes1 = "Апрель"; break;
            //    case "5": mes1 = "Май"; break;
            //    case "6": mes1 = "Июнь"; break;
            //    case "7": mes1 = "Июль"; break;
            //    case "8": mes1 = "Август"; break;
            //    case "9": mes1 = "Сентябрь"; break;
            //    case "10": mes1 = "Октябрь"; break;
            //    case "11": mes1 = "Ноябрь"; break;
            //    case "12": mes1 = "Декабрь"; break;

            //}

            //switch (mes_k)
            //{
            //    case "1": mes2 = "Январь"; break;
            //    case "2": mes2 = "Февраль"; break;
            //    case "3": mes2 = "Март"; break;
            //    case "4": mes2 = "Апрель"; break;
            //    case "5": mes2 = "Май"; break;
            //    case "6": mes2 = "Июнь"; break;
            //    case "7": mes2 = "Июль"; break;
            //    case "8": mes2 = "Август"; break;
            //    case "9": mes2 = "Сентябрь"; break;
            //    case "10": mes2 = "Октябрь"; break;
            //    case "11": mes2 = "Ноябрь"; break;
            //    case "12": mes2 = "Декабрь"; break;

            //}

            if (cb_base.SelectedItem.ToString() == "Компас РП")
            {

                NpgsqlConnection connection = new NpgsqlConnection(monitor_options);
                NpgsqlConnection connection1 = new NpgsqlConnection(monitor);
                connection1.Open();
                connection.Open();
                NpgsqlCommand com = new NpgsqlCommand("select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and \"vvod\" = TRUE", connection);
                rtb.Text = rtb.Text + "Небаланс период \n \n" + "select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and \"vvod\" = TRUE \n \n";
                NpgsqlDataReader datar = com.ExecuteReader();

                while (datar.Read())
                {

                    list_v.Add(datar.GetValue(0).ToString());
                    list_vn.Add(datar.GetValue(0).ToString());

                }

                com = new NpgsqlCommand("select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and \"fider\" = TRUE ", connection);
                rtb.Text = rtb.Text + "select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and \"fider\" = TRUE \n \n";
                datar = com.ExecuteReader();

                while (datar.Read())
                {
                    list_f.Add(datar.GetValue(0).ToString());
                    list_fn.Add(datar.GetValue(0).ToString());

                }

                string g_id = string.Empty;
                NpgsqlCommand com2 = new NpgsqlCommand("select t1.\"Groups_ID\" from \"GROUPS\" t1, \"GROUPS\" t2 where t1.\"Tag\" = 'ТИИ' and t1.\"Parent\"=t2.\"Groups_ID\" and t2.\"Name\" = '" + cb_rp.SelectedItem.ToString() + "'", connection1);
                //MessageBox.Show("select t1.\"Groups_ID\" from \"GROUPS\" t1, \"GROUPS\" t2 where t1.\"Tag\" = 'ТИИ' and t1.\"Parent\"=t2.\"Groups_ID\" and t2.\"Name\" = '" + form5.comboBox1.SelectedItem.ToString() + "'");
                rtb.Text = rtb.Text + "select t1.\"Groups_ID\" from \"GROUPS\" t1, \"GROUPS\" t2 where t1.\"Tag\" = 'ТИИ' and t1.\"Parent\"=t2.\"Groups_ID\" and t2.\"Name\" = '" + cb_rp.SelectedItem.ToString() + "' \n \n";
                g_id = Convert.ToString(com2.ExecuteScalar());


                for (int i = 0; i < list_v.Count; i++)
                {
                    NpgsqlCommand com3 = new NpgsqlCommand("select \"Values_ID\" from \"TB_VALUES\" where \"Groups_ID\" =" + g_id + " and \"Name\" = '" + list_v[i] + "' ", connection1);
                    //MessageBox.Show("select \"Values_ID\" from \"TB_VALUES\" where \"Groups_ID\" =" + g_id + " and \"Name\" = '" + list_v[i] + "' ");
                    rtb.Text = rtb.Text + "select \"Values_ID\" from \"TB_VALUES\" where \"Groups_ID\" =" + g_id + " and \"Name\" = '" + list_v[i] + "' \n \n";
                    list_v[i] = Convert.ToString(com3.ExecuteScalar());
                }


                for (int i = 0; i < list_f.Count; i++)
                {
                    NpgsqlCommand com4 = new NpgsqlCommand("select \"Values_ID\" from \"TB_VALUES\" where \"Groups_ID\" =" + g_id + " and \"Name\" = '" + list_f[i] + "' ", connection1);
                    rtb.Text = rtb.Text + "select \"Values_ID\" from \"TB_VALUES\" where \"Groups_ID\" =" + g_id + " and \"Name\" = '" + list_f[i] + "' \n \n";
                    list_f[i] = Convert.ToString(com4.ExecuteScalar());
                }

                connection1.Close();
                connection.Close();
            }

            if (cb_base.SelectedItem.ToString() == "Энергомера")
            {
                NpgsqlConnection connection = new NpgsqlConnection(monitor_options);
                NpgsqlConnection connection1 = new NpgsqlConnection(energomera_comobjects);
                connection.Open();
                connection1.Open();
                NpgsqlCommand com = new NpgsqlCommand("select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and \"vvod\" = TRUE", connection);
                rtb.Text = rtb.Text + "Небаланс период \n \n" + "select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and \"vvod\" = TRUE \n \n";
                NpgsqlDataReader datar = com.ExecuteReader();

                while (datar.Read())
                {

                    list_v.Add(datar.GetValue(0).ToString());
                    list_vn.Add(datar.GetValue(0).ToString());

                }

                com = new NpgsqlCommand("select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and \"fider\" = TRUE ", connection);
                rtb.Text = rtb.Text + "select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and \"fider\" = TRUE \n \n";
                datar = com.ExecuteReader();

                while (datar.Read())
                {
                    list_f.Add(datar.GetValue(0).ToString());
                    list_fn.Add(datar.GetValue(0).ToString());

                }

                for (int i = 0; i < list_v.Count; i++)
                {
                    NpgsqlCommand com3 = new NpgsqlCommand("select \"MeterId\" from \"BalanceGroups_Meters\" where \"BalanceGroupId\" IN (select \"BalanceGroupId\" from \"BalanceGroups\" where \"Name\" IN ('" + list_v[i] + "'))", connection1);
                    //MessageBox.Show("select \"Values_ID\" from \"TB_VALUES\" where \"Groups_ID\" =" + g_id + " and \"Name\" = '" + list_v[i] + "' ");
                    rtb.Text = rtb.Text + "select \"MeterId\" from \"BalanceGroups_Meters\" where \"BalanceGroupId\" IN (select \"BalanceGroupId\" from \"BalanceGroups\" where \"Name\" IN ('" + list_v[i] + "')) \n \n";
                    list_v[i] = Convert.ToString(com3.ExecuteScalar());
                }


                for (int i = 0; i < list_f.Count; i++)
                {
                    NpgsqlCommand com4 = new NpgsqlCommand("select \"MeterId\" from \"BalanceGroups_Meters\" where \"BalanceGroupId\" IN (select \"BalanceGroupId\" from \"BalanceGroups\" where \"Name\" IN ('" + list_f[i] + "'))", connection1);
                    rtb.Text = rtb.Text + "select \"MeterId\" from \"BalanceGroups_Meters\" where \"BalanceGroupId\" IN (select \"BalanceGroupId\" from \"BalanceGroups\" where \"Name\" IN ('" + list_f[i] + "')) \n \n";
                    list_f[i] = Convert.ToString(com4.ExecuteScalar());
                }


                connection.Close();
                connection1.Close();

            }

            if (cms.Items[8].Text == "Небаланс РП период")
            {
                b.Clear();
                b1.Clear();
                dat.Clear();

                //for (byte i = 0; i <= dayCount; i++)
                //{

                //dat.Add(new DateTime(dtp_n.Value.Year, dtp_n.Value.Month, dtp_n.Value.Day).AddMonths(0));
                string per = dtp_n.Value.ToShortDateString() + "-" + dtp_k.Value.ToShortDateString();
                decimal sum_v1 = 0;
                decimal sum_f1 = 0;

                for (int k = 0; k < list_v.Count; k++)
                {
                    sum_v1 = sum_v1 + summ_period(list_v[k], list_vn[k], dtp_n, dtp_k, cb_base, rtb);

                }

                for (int k = 0; k < list_f.Count; k++)
                {
                    sum_f1 = sum_f1 + summ_period(list_f[k], list_fn[k], dtp_n, dtp_k, cb_base, rtb);
                }
                //MessageBox.Show(sum_v1.ToString() + " - " + sum_f1.ToString());
                b.Add(Math.Round((sum_v1 - sum_f1), 3));
                try
                {
                    b1.Add(Math.Round(((sum_v1 - sum_f1) / sum_v1) * 100, 3));
                }
                catch { b1.Add(0); }
                //catch { MessageBox.Show("Сумма вводов равна нулю. На ноль делить нельзя. \n Проверьте показания."); } 


                class_nebalans m = new class_nebalans();
                m.data = per;
                m.value = b[0].ToString();
                m.proc = b1[0].ToString();
                list_nebalans.Add(m);



                //}

            }
            dgv.DataSource = list_nebalans;
            dgv.Columns[0].HeaderText = "Дата";
            dgv.Columns[1].HeaderText = "кВт*ч";
            dgv.Columns[2].HeaderText = "%";
            dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
            //int ss = 0;
            //ch.Series[0].Points.Clear();
            //while (ss < b.Count)
            //{

            //    var proc = b1[ss];
            //    var skip = b[ss];
            //    var date = dat[ss];
            //    //chart1.ChartAreas[0].AxisY.MajorGrid.Interval = 50;
            //    ch.ChartAreas[0].AxisX.LabelStyle.Format = "dd";
            //    ch.ChartAreas[0].AxisX.IntervalType = System.Windows.Forms.DataVisualization.Charting.DateTimeIntervalType.Days;
            //    ch.ChartAreas[0].AxisX.Interval = 1;
            //    ch.Titles[0].Text = "График небаланса   " + cb_rp.SelectedItem.ToString() + "   за период   " + d1 + "   -   " + d2 + ru_vis;
            //    foreach (Legend objLegend in ch.Legends) //убрали легенду из графика
            //    {
            //        objLegend.Enabled = false;
            //    }
            //    ch.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
            //    ch.Series[0].Points.AddXY(Convert.ToDateTime(date), Convert.ToDouble(skip));
            //    ch.Series[0].ToolTip = "X=#VALX, Y=#VALY";
            //    ss++;
            //}


            return dgv;
        }
        public DataGridView sut_potreb(DataGridView dgv, DateTimePicker dtp_n, DateTimePicker dtp_k, ComboBox cb_rp, ContextMenuStrip cms, Chart ch, ComboBox cb_base, RichTextBox rtb)
        {
            dgv.DataSource = null;
            list_sut_potreb.Clear();
            list_f.Clear();
            list_v.Clear();
            list_vn.Clear();
            list_fn.Clear();

            cms.Items[6].Enabled = true;
            cms.Items[4].Enabled = false;
            cms.Items[7].Enabled = false;

            ktt(cb_rp);

            string rp = cb_rp.SelectedItem.ToString();
            DateTime d1 = new DateTime(dtp_n.Value.Year, dtp_n.Value.Month, dtp_n.Value.Day); //дата начала периода
            double nd = d1.ToOADate();    //перевод нормальной даты в абсолютную
            DateTime d2 = new DateTime(dtp_k.Value.Year, dtp_k.Value.Month, dtp_k.Value.Day); //дата конца периода
            double kd = d2.ToOADate();    //перевод нормальной даты в абсолютную

            //if (kd - nd == 0) { MessageBox.Show("Выбрать дату"); return; }
            long base_tick = new DateTime(1601, 1, 1).Ticks;
            long n_tick = d1.Ticks;
            long n_tick_mysql = n_tick - base_tick;
            long k_tick = d2.Ticks;
            long k_tick_mysql = k_tick - base_tick;

            //if (kd - nd == 0) { MessageBox.Show("Выбрать дату"); return; }
            long dayCount = k_tick_mysql - n_tick_mysql;
            long oneday = 864000000000;
            //byte dayCount = Convert.ToByte(kd - nd);


            string mes_n = dtp_n.Value.Month.ToString();
            string mes_k = dtp_k.Value.Month.ToString();


            switch (mes_n)
            {
                case "1": mes1 = "Январь"; break;
                case "2": mes1 = "Февраль"; break;
                case "3": mes1 = "Март"; break;
                case "4": mes1 = "Апрель"; break;
                case "5": mes1 = "Май"; break;
                case "6": mes1 = "Июнь"; break;
                case "7": mes1 = "Июль"; break;
                case "8": mes1 = "Август"; break;
                case "9": mes1 = "Сентябрь"; break;
                case "10": mes1 = "Октябрь"; break;
                case "11": mes1 = "Ноябрь"; break;
                case "12": mes1 = "Декабрь"; break;

            }

            switch (mes_k)
            {
                case "1": mes2 = "Январь"; break;
                case "2": mes2 = "Февраль"; break;
                case "3": mes2 = "Март"; break;
                case "4": mes2 = "Апрель"; break;
                case "5": mes2 = "Май"; break;
                case "6": mes2 = "Июнь"; break;
                case "7": mes2 = "Июль"; break;
                case "8": mes2 = "Август"; break;
                case "9": mes2 = "Сентябрь"; break;
                case "10": mes2 = "Октябрь"; break;
                case "11": mes2 = "Ноябрь"; break;
                case "12": mes2 = "Декабрь"; break;

            }
            if (cb_base.SelectedItem.ToString() == "Компас РП")
            {
                NpgsqlConnection connection = new NpgsqlConnection(monitor_options);
                NpgsqlConnection connection1 = new NpgsqlConnection(monitor);
                connection1.Open();
                connection.Open();
                NpgsqlCommand com = new NpgsqlCommand("select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and (\"vvod\" = TRUE or \"vvod_niz\" = TRUE)", connection);
                rtb.Text = rtb.Text + "Сут потреб \n \n" + "select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and (\"vvod\" = TRUE or \"vvod_niz\" = TRUE)\n \n";
                NpgsqlDataReader datar = com.ExecuteReader();

                while (datar.Read())
                {

                    list_v.Add(datar.GetValue(0).ToString());
                    list_vn.Add(datar.GetValue(0).ToString());

                }

                com = new NpgsqlCommand("select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and (\"fider\" = TRUE or \"fider_niz\" = TRUE)", connection);
                rtb.Text = rtb.Text + "select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and (\"fider\" = TRUE or \"fider_niz\" = TRUE) \n \n";
                datar = com.ExecuteReader();

                while (datar.Read())
                {
                    list_f.Add(datar.GetValue(0).ToString());
                    list_fn.Add(datar.GetValue(0).ToString());

                }

                string g_id = string.Empty;
                NpgsqlCommand com2 = new NpgsqlCommand("select t1.\"Groups_ID\" from \"GROUPS\" t1, \"GROUPS\" t2 where t1.\"Tag\" = 'ТИИ' and t1.\"Parent\"=t2.\"Groups_ID\" and t2.\"Name\" = '" + cb_rp.SelectedItem.ToString() + "'", connection1);
                //MessageBox.Show("select t1.\"Groups_ID\" from \"GROUPS\" t1, \"GROUPS\" t2 where t1.\"Tag\" = 'ТИИ' and t1.\"Parent\"=t2.\"Groups_ID\" and t2.\"Name\" = '" + form5.comboBox1.SelectedItem.ToString() + "'");
                rtb.Text = rtb.Text + "select t1.\"Groups_ID\" from \"GROUPS\" t1, \"GROUPS\" t2 where t1.\"Tag\" = 'ТИИ' and t1.\"Parent\"=t2.\"Groups_ID\" and t2.\"Name\" = '" + cb_rp.SelectedItem.ToString() + "' \n \n";
                g_id = Convert.ToString(com2.ExecuteScalar());


                for (int i = 0; i < list_v.Count; i++)
                {
                    NpgsqlCommand com3 = new NpgsqlCommand("select \"Values_ID\" from \"TB_VALUES\" where \"Groups_ID\" =" + g_id + " and \"Name\" = '" + list_v[i] + "' ", connection1);
                    //MessageBox.Show("select \"Values_ID\" from \"TB_VALUES\" where \"Groups_ID\" =" + g_id + " and \"Name\" = '" + list_v[i] + "' ");
                    rtb.Text = rtb.Text + "select \"Values_ID\" from \"TB_VALUES\" where \"Groups_ID\" =" + g_id + " and \"Name\" = '" + list_v[i] + "' ";
                    list_v[i] = Convert.ToString(com3.ExecuteScalar());
                }


                for (int i = 0; i < list_f.Count; i++)
                {
                    NpgsqlCommand com4 = new NpgsqlCommand("select \"Values_ID\" from \"TB_VALUES\" where \"Groups_ID\" =" + g_id + " and \"Name\" = '" + list_f[i] + "' ", connection1);
                    rtb.Text = rtb.Text + "select \"Values_ID\" from \"TB_VALUES\" where \"Groups_ID\" =" + g_id + " and \"Name\" = '" + list_f[i] + "' ";
                    list_f[i] = Convert.ToString(com4.ExecuteScalar());
                }

                connection1.Close();
                connection.Close();
            }

            if (cb_base.SelectedItem.ToString() == "Энергомера")
            {
                NpgsqlConnection connection12 = new NpgsqlConnection(monitor_options);
                NpgsqlConnection connection11 = new NpgsqlConnection(energomera_comobjects);
                connection12.Open();
                connection11.Open();
                NpgsqlCommand com12 = new NpgsqlCommand("select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and \"vvod\" = TRUE", connection12);
                rtb.Text = rtb.Text + "Сут потреб \n \n" + "select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and \"vvod\" = TRUE \n \n";
                NpgsqlDataReader datar12 = com12.ExecuteReader();

                while (datar12.Read())
                {

                    list_v.Add(datar12.GetValue(0).ToString());
                    list_vn.Add(datar12.GetValue(0).ToString());

                }

                com12 = new NpgsqlCommand("select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and \"fider\" = TRUE ", connection12);
                rtb.Text = rtb.Text + "select \"name_tii\"  from \"balans\" where \"name_rp\" = '" + rp + "' and \"fider\" = TRUE \n \n";
                datar12 = com12.ExecuteReader();

                while (datar12.Read())
                {
                    list_f.Add(datar12.GetValue(0).ToString());
                    list_fn.Add(datar12.GetValue(0).ToString());

                }

                for (int i = 0; i < list_v.Count; i++)
                {
                    NpgsqlCommand com3 = new NpgsqlCommand("select \"MeterId\" from \"BalanceGroups_Meters\" where \"BalanceGroupId\" IN (select \"BalanceGroupId\" from \"BalanceGroups\" where \"Name\" IN ('" + list_v[i] + "'))", connection11);
                    //MessageBox.Show("select \"Values_ID\" from \"TB_VALUES\" where \"Groups_ID\" =" + g_id + " and \"Name\" = '" + list_v[i] + "' ");
                    rtb.Text = rtb.Text + "select \"MeterId\" from \"BalanceGroups_Meters\" where \"BalanceGroupId\" IN (select \"BalanceGroupId\" from \"BalanceGroups\" where \"Name\" IN ('" + list_v[i] + "')) \n \n";
                    list_v[i] = Convert.ToString(com3.ExecuteScalar());
                }


                for (int i = 0; i < list_f.Count; i++)
                {
                    NpgsqlCommand com4 = new NpgsqlCommand("select \"MeterId\" from \"BalanceGroups_Meters\" where \"BalanceGroupId\" IN (select \"BalanceGroupId\" from \"BalanceGroups\" where \"Name\" IN ('" + list_f[i] + "'))", connection11);
                    rtb.Text = rtb.Text + "select \"MeterId\" from \"BalanceGroups_Meters\" where \"BalanceGroupId\" IN (select \"BalanceGroupId\" from \"BalanceGroups\" where \"Name\" IN ('" + list_f[i] + "')) \n \n";
                    list_f[i] = Convert.ToString(com4.ExecuteScalar());
                }


                connection12.Close();
                connection11.Close();

            }

            if (cms.Items[5].Text == "Суточное потребление")
            {
                b.Clear();
                b1.Clear();
                dat.Clear();

               // for (byte i = 0; i <= dayCount; i++)
                    for (long i = 0; i <= dayCount; i = i + oneday)     
                {
                    //MessageBox.Show(i.ToString() + "--" + dayCount);
                    long x1 = Convert.ToInt64(i + base_tick);
                    DateTime dt1 = new DateTime(x1);
                    //dat.Add(new DateTime(dtp_n.Value.Year, dtp_n.Value.Month, dtp_n.Value.Day).AddDays(i));
                    dat.Add(new DateTime(dtp_n.Value.Year, dtp_n.Value.Month, dtp_n.Value.Day).AddDays(dt1.Day - 1));

                    decimal sum_v1 = 0;
                    decimal sum_f1 = 0;

                    for (int k = 0; k < list_v.Count; k++)
                    {
                        sum_v1 = sum_v1 + summ(list_v[k], i, list_vn[k], dtp_n, cb_base, rtb);

                    }

                    for (int k = 0; k < list_f.Count; k++)
                    {
                        sum_f1 = sum_f1 + summ(list_f[k], i, list_fn[k], dtp_n, cb_base, rtb);
                    }
                    //MessageBox.Show(sum_v1.ToString() + " - " + sum_f1.ToString());
                    b.Add(Math.Round(sum_f1, 3));


                    b1.Add(Math.Round(sum_v1, 3));


                    class_sut_potreb m = new class_sut_potreb();
                    m.data = dat[dt1.Day - 1].ToString();
                    m.value = b[dt1.Day - 1].ToString();
                    m.proc = b1[dt1.Day - 1].ToString();
                    list_sut_potreb.Add(m);
                }
            }
            dgv.DataSource = list_sut_potreb;
            dgv.Columns[0].HeaderText = "Дата";
            dgv.Columns[1].HeaderText = "Энергия(кВт*ч)";
            dgv.Columns[2].HeaderText = "За сутки(кВт*ч)";
            dgv.Columns[3].HeaderText = "На вводах(кВт*ч)";

            for (int k = 0; k < dgv.RowCount; k++)
            {
                if (k == 0) { dgv.Rows[k].Cells[1].Value = dgv.Rows[k].Cells[2].Value; }
                if (k > 0) { dgv.Rows[k].Cells[1].Value = Convert.ToDecimal(dgv.Rows[k - 1].Cells[1].Value) + Convert.ToDecimal(dgv.Rows[k].Cells[2].Value); }
            }

            dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);

            int ss = 0;
            ch.Series[0].Points.Clear();
            while (ss < b.Count)
            {


                var skip = b[ss];
                var date = dat[ss];
                //chart1.ChartAreas[0].AxisY.MajorGrid.Interval = 50;
                ch.ChartAreas[0].AxisX.LabelStyle.Format = "dd";
                ch.ChartAreas[0].AxisX.IntervalType = System.Windows.Forms.DataVisualization.Charting.DateTimeIntervalType.Days;
                ch.ChartAreas[0].AxisX.Interval = 1;
                ch.Titles[0].Text = "График суточного потребления " + cb_rp.SelectedItem.ToString() + "  за период  " + d1 + "   -   " + d2;
                foreach (Legend objLegend in ch.Legends) //убрали легенду из графика
                {
                    objLegend.Enabled = false;
                }
                ch.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                ch.Series[0].Points.AddXY(Convert.ToDateTime(date), Convert.ToDouble(skip));
                ch.Series[0].ToolTip = "X=#VALX, Y=#VALY";
                ss++;
            }
            return dgv;
        } //суточное потребление + график
        public void nebalans_sutki(DataGridView dgv, DataGridView dgv_vrem, CheckBox cb_niz, ComboBox cb_rp, DateTimePicker dtp, RichTextBox rtb)
        {
            prov = 1;
            dgv.DataSource = null;
            dgv_vrem.DataSource = null;
            dgv.Rows.Clear();
            dgv_vrem.Rows.Clear();
            list1.Clear();



            string vv = "";
            string ff = "";

            if (cb_niz.Checked == true) { vv = "vvod_niz"; ff = "fider_niz"; }
            if (cb_niz.Checked == false) { vv = "vvod"; ff = "fider"; }

            //MessageBox.Show(v_id.ToString());
            DateTime d1 = new DateTime(dtp.Value.Year, dtp.Value.Month, dtp.Value.Day); //дата начала периода
            double x1 = d1.ToOADate();    //перевод нормальной даты в абсолютную


            long base_tick = new DateTime(1601, 1, 1).Ticks;
            long n_tick = d1.Ticks;
            long n_tick_mysql = n_tick - base_tick;
            


            string chas1 = "0:00";
            chas1 = chas1.Substring(0, chas1.IndexOf(":"));
            string min1 = "0:00";
            min1 = min1.Substring(min1.IndexOf(":") + 1, 2);
            double ab_vr1 = (Convert.ToDouble(chas1) * 60 + Convert.ToDouble(min1)) * 0.000694443; //перевод часов и минут (то есть время) в абсолютное. Это от

            long tick_vr1 = (Convert.ToInt64(chas1) * 60 + Convert.ToInt64(min1))*600000000;

            string chas2 = "23:30";
            chas2 = chas2.Substring(0, chas2.IndexOf(":"));
            string min2 = "23:30";
            min2 = min2.Substring(min2.IndexOf(":") + 1, 2);
            double ab_vr2 = (Convert.ToDouble(chas2) * 60 + 2 * Convert.ToDouble(min2)) * 0.000694445;  //перевод часов и минут (то есть время) в абсолютное. Это до

            long tick_vr2 = (Convert.ToInt64(chas2) * 60 + 2 * Convert.ToInt64(min2))*600000000;

           

            double vr1 = x1 + ab_vr1;
            double vr3 = x1 + ab_vr2;

            long vr1_mysql = n_tick_mysql + tick_vr1;
            long vr3_mysql = n_tick_mysql + tick_vr2;

           // MessageBox.Show(vr1_mysql.ToString() + "--" + vr3_mysql.ToString());

            string vr11 = Convert.ToString(vr1).Replace(",", ".");
            string vr22 = Convert.ToString(vr3).Replace(",", ".");

            string n_d = DateTime.FromOADate(vr1).ToString().Replace(".", "-");
            string k_d = DateTime.FromOADate(vr3).ToString().Replace(".", "-");

            NpgsqlConnection connection1 = new NpgsqlConnection(monitor_options);
            connection1.Open();
            NpgsqlConnection connection = new NpgsqlConnection(monitor);
            NpgsqlDataAdapter adapter = new NpgsqlDataAdapter();

            connection.Open();
            //if (cb_rp.SelectedItem == null) { MessageBox.Show("Выбрать РП"); return; }
            NpgsqlCommand command = new NpgsqlCommand("select t4.\"Values_ID\", t0.\"Name\", t4.\"Name\", t4.\"Tag\" from (select t1.\"Groups_ID\", t2.\"Name\" from \"GROUPS\" t1, \"GROUPS\" t2 where t1.\"Tag\" = 'ТИИ' and t1.\"Parent\"=t2.\"Groups_ID\" and t2.\"Name\" = '" + cb_rp.SelectedItem.ToString() + "') as t0, \"TB_VALUES\" t4 where t4.\"Groups_ID\"=t0.\"Groups_ID\" ", connection);
            rtb.Text = rtb.Text + "Небаланс сутки \n \n" + "select t4.\"Values_ID\", t0.\"Name\", t4.\"Name\", t4.\"Tag\" from (select t1.\"Groups_ID\", t2.\"Name\" from \"GROUPS\" t1, \"GROUPS\" t2 where t1.\"Tag\" = 'ТИИ' and t1.\"Parent\"=t2.\"Groups_ID\" and t2.\"Name\" = '" + cb_rp.SelectedItem.ToString() + "') as t0, \"TB_VALUES\" t4 where t4.\"Groups_ID\"=t0.\"Groups_ID\" \n \n";
            adapter.SelectCommand = command;
            DataTable table = new DataTable();
            adapter.Fill(table);


            foreach (DataRow row in table.Rows)
            {
                class2 m = new class2();
                m.values_id = row.ItemArray[0].ToString();
                m.name = row.ItemArray[1].ToString();
                m.name_f = row.ItemArray[2].ToString();
                m.tag = row.ItemArray[3].ToString();
                list1.Add(m);
            }
            string rp = cb_rp.SelectedItem.ToString();
            List<string> list_f = new List<string>(); //лист фидеров
            List<string> list_v = new List<string>(); //лист вводов
            NpgsqlCommand com1 = new NpgsqlCommand("select \"name_tii\", \"ktt\"  from balans where \"name_rp\" = '" + rp + "' and \"" + ff + "\" = TRUE", connection1);
            rtb.Text = rtb.Text + "select \"name_tii\", \"ktt\"  from balans where \"name_rp\" = '" + rp + "' and \"" + ff + "\" = TRUE \n \n";
            NpgsqlDataReader datar1 = com1.ExecuteReader();

            fider_koef.Clear();

            while (datar1.Read())
            {
                class8 m = new class8();
                m.name_f = datar1.GetValue(0).ToString();
                m.koef = datar1.GetValue(1).ToString();
                fider_koef.Add(m);

            }




            com1 = new NpgsqlCommand("select \"name_tii\", \"ktt\"  from balans where \"name_rp\" = '" + rp + "' and \"" + vv + "\" = TRUE", connection1);
            rtb.Text = rtb.Text + "select \"name_tii\", \"ktt\"  from balans where \"name_rp\" = '" + rp + "' and \"" + vv + "\" = TRUE \n \n";
            datar1 = com1.ExecuteReader();
            vvod_koef.Clear();
            while (datar1.Read())
            {
                class8_1 m = new class8_1();
                m.name_f = datar1.GetValue(0).ToString();
                m.koef = datar1.GetValue(1).ToString();
                vvod_koef.Add(m);

            }


            string g_id = string.Empty;
            command = new NpgsqlCommand("select t1.\"Groups_ID\" from \"GROUPS\" t1, \"GROUPS\" t2 where t1.\"Tag\" = 'ТИИ' and t1.\"Parent\"=t2.\"Groups_ID\" and t2.\"Name\" = '" + cb_rp.SelectedItem.ToString() + "'", connection);
            rtb.Text = rtb.Text + "select t1.\"Groups_ID\" from \"GROUPS\" t1, \"GROUPS\" t2 where t1.\"Tag\" = 'ТИИ' and t1.\"Parent\"=t2.\"Groups_ID\" and t2.\"Name\" = '" + cb_rp.SelectedItem.ToString() + "' \n \n";
            g_id = Convert.ToString(command.ExecuteScalar());

            for (int i = 0; i < fider_koef.Count; i++)
            {

                command = new NpgsqlCommand("select \"Values_ID\" from \"TB_VALUES\" where \"Groups_ID\" =" + g_id + " and \"Name\" = '" + fider_koef[i].name_f + "' ", connection);
                rtb.Text = rtb.Text + "select \"Values_ID\" from \"TB_VALUES\" where \"Groups_ID\" =" + g_id + " and \"Name\" = '" + fider_koef[i].name_f + "' \n \n";
                list_f.Add(Convert.ToString(command.ExecuteScalar()));

            }

            for (int i = 0; i < vvod_koef.Count; i++)
            {

                command = new NpgsqlCommand("select \"Values_ID\" from \"TB_VALUES\" where \"Groups_ID\" =" + g_id + " and \"Name\" = '" + vvod_koef[i].name_f + "' ", connection);
                rtb.Text = rtb.Text + "select \"Values_ID\" from \"TB_VALUES\" where \"Groups_ID\" =" + g_id + " and \"Name\" = '" + vvod_koef[i].name_f + "' \n \n";
                list_v.Add(Convert.ToString(command.ExecuteScalar()));

            }

            connection1.Close();
            connection.Close();

            StringBuilder v_id = new StringBuilder();
            v_id.Append("(");
            for (int i = 0; i < list_f.Count; i++)
            {
                v_id.Append(list_f[i]);

                if (i != list_f.Count - 1)
                {
                    v_id.Append(", ");
                }
            }
            v_id.Append(")");

            StringBuilder v1_id = new StringBuilder();
            v1_id.Append("(");
            for (int i = 0; i < list_v.Count; i++)
            {
                v1_id.Append(list_v[i]);

                if (i != list_v.Count - 1)
                {
                    v1_id.Append(", ");
                }
            }
            v1_id.Append(")");




           // SqlConnection con = new SqlConnection(sql_connection);
            MySqlConnection con = new MySqlConnection(mysql_connection);
            con.Open();
            //SqlDataAdapter da = new SqlDataAdapter("select VALUES_ID, T_TIME,  V_VALUE, VALID from ti where T_TIME >= " + vr11 + " and T_TIME <= " + vr22 + " and values_id in " + v_id.ToString() + " order by VALUES_ID, T_TIME ", con);
            //rtb.Text = rtb.Text + "select VALUES_ID, T_TIME,  V_VALUE, VALID from ti where T_TIME >= " + vr11 + " and T_TIME <= " + vr22 + " and values_id in " + v_id.ToString() + " order by VALUES_ID, T_TIME \n \n";
            MySqlDataAdapter da = new MySqlDataAdapter("select VALUES_ID, T_TIME,  V_VALUE, VALID from ti where T_TIME >= " + vr1_mysql + " and T_TIME <= " + vr3_mysql + " and values_id in " + v_id.ToString() + " order by VALUES_ID, T_TIME ", con);
            rtb.Text = rtb.Text + "select VALUES_ID, T_TIME,  V_VALUE, VALID from ti where T_TIME >= " + vr1_mysql + " and T_TIME <= " + vr3_mysql + " and values_id in " + v_id.ToString() + " order by VALUES_ID, T_TIME \n \n";
            //SqlCommandBuilder cb = new SqlCommandBuilder(da);
            MySqlCommandBuilder cb = new MySqlCommandBuilder(da);
            DataTable table1 = new DataTable();
            da.Fill(table1);
            list.Clear();
            foreach (DataRow row in table1.Rows)
            {
                class0 m = new class0();
                m.t_time = row.ItemArray[1].ToString();
                m.values_id = row.ItemArray[0].ToString();
                m.v_value = row.ItemArray[2].ToString();
                m.valid = row.ItemArray[3].ToString();
                list.Add(m);
            }
            con.Close();


            //SqlConnection con1 = new SqlConnection(sql_connection);
            MySqlConnection con1 = new MySqlConnection(mysql_connection);
            con1.Open();
            //SqlDataAdapter da1 = new SqlDataAdapter("select VALUES_ID, T_TIME,  V_VALUE, VALID from ti where T_TIME >= " + vr11 + " and T_TIME <= " + vr22 + " and values_id in " + v1_id.ToString() + " order by VALUES_ID, T_TIME ", con);
            //rtb.Text = rtb.Text + "select VALUES_ID, T_TIME,  V_VALUE, VALID from ti where T_TIME >= " + vr11 + " and T_TIME <= " + vr22 + " and values_id in " + v1_id.ToString() + " order by VALUES_ID, T_TIME  \n \n";
            MySqlDataAdapter da1 = new MySqlDataAdapter("select VALUES_ID, T_TIME,  V_VALUE, VALID from ti where T_TIME >= " + vr1_mysql + " and T_TIME <= " + vr3_mysql + " and values_id in " + v1_id.ToString() + " order by VALUES_ID, T_TIME ", con);
            rtb.Text = rtb.Text + "select VALUES_ID, T_TIME,  V_VALUE, VALID from ti where T_TIME >= " + vr1_mysql + " and T_TIME <= " + vr3_mysql + " and values_id in " + v1_id.ToString() + " order by VALUES_ID, T_TIME  \n \n";
            //SqlCommandBuilder cb1 = new SqlCommandBuilder(da1);
            MySqlCommandBuilder cb1 = new MySqlCommandBuilder(da1);
            DataTable table2 = new DataTable();
            da1.Fill(table2);
            list_vlist.Clear();
            foreach (DataRow row in table2.Rows)
            {
                class9 m = new class9();
                m.t_time = row.ItemArray[1].ToString();
                m.values_id = row.ItemArray[0].ToString();
                m.v_value = row.ItemArray[2].ToString();
                m.valid = row.ItemArray[3].ToString();
                list_vlist.Add(m);
            }
            con1.Close();



            for (int i = 0; i < list.Count; i++)
            {
                //double x = Convert.ToDouble(list[i].t_time);
                //DateTime dt = DateTime.FromOADate(x);

                long x_mysql = Convert.ToInt64(list[i].t_time) + base_tick;
                DateTime dt = new DateTime(x_mysql);
                list[i].t_time = dt.ToString();
            }

            for (int i = 0; i < list_vlist.Count; i++)
            {
                //double x = Convert.ToDouble(list_vlist[i].t_time);
                //DateTime dt = DateTime.FromOADate(x);

                long x_mysql = Convert.ToInt64(list_vlist[i].t_time) + base_tick;
                DateTime dt = new DateTime(x_mysql);

                list_vlist[i].t_time = dt.ToString();
            }


            result.Clear();
            result_v.Clear();

            for (int i = 0; i < list.Count; i++)
            {
                for (int j = 0; j < list1.Count; j++)
                {
                    if (list[i].values_id == list1[j].values_id)
                    {
                        class3 m = new class3();
                        m.name = list1[j].name;
                        m.name_f = list1[j].name_f;
                        m.tag = list1[j].tag;
                        m.t_time = list[i].t_time;
                        m.v_value = list[i].v_value;
                        m.valid = list[i].valid;
                        result.Add(m);
                    }
                }
            }

            for (int i = 0; i < list_vlist.Count; i++)
            {
                for (int j = 0; j < list1.Count; j++)
                {
                    if (list_vlist[i].values_id == list1[j].values_id)
                    {
                        class3_1 m = new class3_1();
                        m.name = list1[j].name;
                        m.name_f = list1[j].name_f;
                        m.tag = list1[j].tag;
                        m.t_time = list_vlist[i].t_time;
                        m.v_value = list_vlist[i].v_value;
                        m.valid = list_vlist[i].valid;
                        result_v.Add(m);
                    }
                }
            }
            dgv.DataSource = result;

            asc.Clear();
            asc_v.Clear();
            List<String> ls1 = new List<String>();
            List<String> ls2 = new List<String>();

            for (int i = 0; i < result.Count - 1; i++)
            {
                if (ls1.IndexOf(result[i].name_f) < 0)
                {
                    ls1.Add(result[i].name_f);
                    class4 m = new class4();
                    m.name_rp = result[i].name_f;
                    asc.Add(m);
                }
            }

            for (int k = 0; k < result.Count; k++)
            {

                for (int j = 0; j < asc.Count; j++)
                {

                    if (result[k].name_f.Equals(asc[j].name_rp))
                    {

                        string streng = result[k].t_time.Substring(result[k].t_time.IndexOf(" "), result[k].t_time.Length - result[k].t_time.IndexOf(" ")).Trim();
                        asc[j].v(timeparser(streng), result[k].v_value);
                    }
                }
            }



            for (int i = 0; i < result_v.Count - 1; i++)
            {
                if (ls2.IndexOf(result_v[i].name_f) < 0)
                {
                    ls2.Add(result_v[i].name_f);
                    class4_1 m = new class4_1();
                    m.name_rp = result_v[i].name_f;
                    asc_v.Add(m);
                }
            }

            for (int k = 0; k < result_v.Count; k++)
            {

                for (int j = 0; j < asc_v.Count; j++)
                {

                    if (result_v[k].name_f.Equals(asc_v[j].name_rp))
                    {

                        string streng1 = result_v[k].t_time.Substring(result_v[k].t_time.IndexOf(" "), result_v[k].t_time.Length - result_v[k].t_time.IndexOf(" ")).Trim();
                        asc_v[j].v(timeparser(streng1), result_v[k].v_value);
                    }
                }
            }


            dgv.DataSource = null;
            dgv.DataSource = asc;
            dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
            dgv.Columns[0].HeaderText = "Название";  //присваиваем текст к заголовку колонок
            dgv.Columns[1].HeaderText = "00:00";
            dgv.Columns[2].HeaderText = "00:30";
            dgv.Columns[3].HeaderText = "01:00";
            dgv.Columns[4].HeaderText = "01:30";
            dgv.Columns[5].HeaderText = "02:00";
            dgv.Columns[6].HeaderText = "02:30";
            dgv.Columns[7].HeaderText = "03:00";
            dgv.Columns[8].HeaderText = "03:30";
            dgv.Columns[9].HeaderText = "04:00";
            dgv.Columns[10].HeaderText = "04:30";
            dgv.Columns[11].HeaderText = "05:00";
            dgv.Columns[12].HeaderText = "05:30";
            dgv.Columns[13].HeaderText = "06:00";
            dgv.Columns[14].HeaderText = "06:30";
            dgv.Columns[15].HeaderText = "07:00";
            dgv.Columns[16].HeaderText = "07:30";
            dgv.Columns[17].HeaderText = "08:00";
            dgv.Columns[18].HeaderText = "08:30";
            dgv.Columns[19].HeaderText = "09:00";
            dgv.Columns[20].HeaderText = "09:30";
            dgv.Columns[21].HeaderText = "10:00";
            dgv.Columns[22].HeaderText = "10:30";
            dgv.Columns[23].HeaderText = "11:00";
            dgv.Columns[24].HeaderText = "11:30";
            dgv.Columns[25].HeaderText = "12:00";
            dgv.Columns[26].HeaderText = "12:30";
            dgv.Columns[27].HeaderText = "13:00";
            dgv.Columns[28].HeaderText = "13:30";
            dgv.Columns[29].HeaderText = "14:00";
            dgv.Columns[30].HeaderText = "14:30";
            dgv.Columns[31].HeaderText = "15:00";
            dgv.Columns[32].HeaderText = "15:30";
            dgv.Columns[33].HeaderText = "16:00";
            dgv.Columns[34].HeaderText = "16:30";
            dgv.Columns[35].HeaderText = "17:00";
            dgv.Columns[36].HeaderText = "17:30";
            dgv.Columns[37].HeaderText = "18:00";
            dgv.Columns[38].HeaderText = "18:30";
            dgv.Columns[39].HeaderText = "19:00";
            dgv.Columns[40].HeaderText = "19:30";
            dgv.Columns[41].HeaderText = "20:00";
            dgv.Columns[42].HeaderText = "20:30";
            dgv.Columns[43].HeaderText = "21:00";
            dgv.Columns[44].HeaderText = "21:30";
            dgv.Columns[45].HeaderText = "22:00";
            dgv.Columns[46].HeaderText = "22:30";
            dgv.Columns[47].HeaderText = "23:00";
            dgv.Columns[48].HeaderText = "23:30";
            dgv.Columns[49].HeaderText = "24:00";


            dgv_vrem.DataSource = null;
            dgv_vrem.DataSource = asc_v;
            dgv_vrem.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
            dgv_vrem.Columns[0].HeaderText = "Название";  //присваиваем текст к заголовку колонок
            dgv_vrem.Columns[1].HeaderText = "00:00";
            dgv_vrem.Columns[2].HeaderText = "00:30";
            dgv_vrem.Columns[3].HeaderText = "01:00";
            dgv_vrem.Columns[4].HeaderText = "01:30";
            dgv_vrem.Columns[5].HeaderText = "02:00";
            dgv_vrem.Columns[6].HeaderText = "02:30";
            dgv_vrem.Columns[7].HeaderText = "03:00";
            dgv_vrem.Columns[8].HeaderText = "03:30";
            dgv_vrem.Columns[9].HeaderText = "04:00";
            dgv_vrem.Columns[10].HeaderText = "04:30";
            dgv_vrem.Columns[11].HeaderText = "05:00";
            dgv_vrem.Columns[12].HeaderText = "05:30";
            dgv_vrem.Columns[13].HeaderText = "06:00";
            dgv_vrem.Columns[14].HeaderText = "06:30";
            dgv_vrem.Columns[15].HeaderText = "07:00";
            dgv_vrem.Columns[16].HeaderText = "07:30";
            dgv_vrem.Columns[17].HeaderText = "08:00";
            dgv_vrem.Columns[18].HeaderText = "08:30";
            dgv_vrem.Columns[19].HeaderText = "09:00";
            dgv_vrem.Columns[20].HeaderText = "09:30";
            dgv_vrem.Columns[21].HeaderText = "10:00";
            dgv_vrem.Columns[22].HeaderText = "10:30";
            dgv_vrem.Columns[23].HeaderText = "11:00";
            dgv_vrem.Columns[24].HeaderText = "11:30";
            dgv_vrem.Columns[25].HeaderText = "12:00";
            dgv_vrem.Columns[26].HeaderText = "12:30";
            dgv_vrem.Columns[27].HeaderText = "13:00";
            dgv_vrem.Columns[28].HeaderText = "13:30";
            dgv_vrem.Columns[29].HeaderText = "14:00";
            dgv_vrem.Columns[30].HeaderText = "14:30";
            dgv_vrem.Columns[31].HeaderText = "15:00";
            dgv_vrem.Columns[32].HeaderText = "15:30";
            dgv_vrem.Columns[33].HeaderText = "16:00";
            dgv_vrem.Columns[34].HeaderText = "16:30";
            dgv_vrem.Columns[35].HeaderText = "17:00";
            dgv_vrem.Columns[36].HeaderText = "17:30";
            dgv_vrem.Columns[37].HeaderText = "18:00";
            dgv_vrem.Columns[38].HeaderText = "18:30";
            dgv_vrem.Columns[39].HeaderText = "19:00";
            dgv_vrem.Columns[40].HeaderText = "19:30";
            dgv_vrem.Columns[41].HeaderText = "20:00";
            dgv_vrem.Columns[42].HeaderText = "20:30";
            dgv_vrem.Columns[43].HeaderText = "21:00";
            dgv_vrem.Columns[44].HeaderText = "21:30";
            dgv_vrem.Columns[45].HeaderText = "22:00";
            dgv_vrem.Columns[46].HeaderText = "22:30";
            dgv_vrem.Columns[47].HeaderText = "23:00";
            dgv_vrem.Columns[48].HeaderText = "23:30";
            dgv_vrem.Columns[49].HeaderText = "24:00";
            //return dgv;
            //return dgv_vrem;

        } //загрузка данных для суточного небаланса в виде широкой таблицы
        public void nebalans_sutki1(DataGridView dgv, DataGridView dgv_vrem)
        {
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                for (int j = 1; j < dgv.Columns.Count - 1; j++)
                {
                    dgv.Rows[i].Cells[j].Value = Convert.ToDecimal(dgv.Rows[i].Cells[j + 1].Value) - Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                    //MessageBox.Show(dgv_vrem.Rows[i].Cells[j].Value.ToString());
                }
            }

            for (int i = 0; i < dgv_vrem.Rows.Count; i++)
            {
                for (int j = 1; j < dgv_vrem.Columns.Count - 1; j++)
                {
                    dgv_vrem.Rows[i].Cells[j].Value = Convert.ToDecimal(dgv_vrem.Rows[i].Cells[j + 1].Value) - Convert.ToDecimal(dgv_vrem.Rows[i].Cells[j].Value);
                }
            }
            //return dgv;
            //return dgv_vrem;
        } //разница показаний по столбцам
        public void nebalans_sutki1_energo(DataGridView dgv, DataGridView dgv_vrem)
        {
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                for (int j = 1; j < dgv.Columns.Count - 1; j++)
                {
                    try
                    {
                        dgv.Rows[i].Cells[j].Value = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value) / 2;
                    }
                    catch { dgv.Rows[i].Cells[j].Value = 0; }
                }
            }

            for (int i = 0; i < dgv_vrem.Rows.Count; i++)
            {
                for (int j = 1; j < dgv_vrem.Columns.Count - 1; j++)
                {
                    try
                    {
                        dgv_vrem.Rows[i].Cells[j].Value = Convert.ToDecimal(dgv_vrem.Rows[i].Cells[j].Value) / 2;
                    }
                    catch { dgv_vrem.Rows[i].Cells[j].Value = 0; }
                }
            }
            //return dgv;
            //return dgv_vrem;
        } //деление данных на два - для энергомеры
        public void nebalans_sutki_koef(DataGridView dgv, DataGridView dgv_vrem)
        {
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                for (int j = 1; j < dgv.Columns.Count - 1; j++)
                {
                    for (int k = 0; k < fider_koef.Count; k++)
                    {
                        if (dgv.Rows[i].Cells[0].Value.ToString().Equals(fider_koef[k].name_f))
                        {

                            dgv.Rows[i].Cells[j].Value = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value) * Convert.ToDecimal(fider_koef[k].koef);

                        }
                    }
                }
            }


            for (int i = 0; i < dgv_vrem.Rows.Count; i++)
            {
                for (int j = 1; j < dgv_vrem.Columns.Count - 1; j++)
                {
                    for (int k = 0; k < vvod_koef.Count; k++)
                    {
                        if (dgv_vrem.Rows[i].Cells[0].Value.ToString().Equals(vvod_koef[k].name_f))
                        {

                            dgv_vrem.Rows[i].Cells[j].Value = Convert.ToDecimal(dgv_vrem.Rows[i].Cells[j].Value) * Convert.ToDecimal(vvod_koef[k].koef);

                        }
                    }
                }
            }
        } //умножение данных на коэффициент
        public void nebalans_sutki_sum_columns(DataGridView dgv, DataGridView dgv_vrem)
        {
            for (int i = 1; i < dgv.Rows.Count; i++)
            {
                for (int j = 1; j < dgv.Columns.Count - 1; j++)
                {
                    dgv.Rows[0].Cells[j].Value = Math.Round(Convert.ToDecimal(dgv.Rows[0].Cells[j].Value) + Convert.ToDecimal(dgv.Rows[i].Cells[j].Value), 3);
                    //MessageBox.Show(dgv.Rows[0].Cells[j].Value.ToString());
                }
            }

            for (int i = 1; i < dgv_vrem.Rows.Count; i++)
            {
                for (int j = 1; j < dgv_vrem.Columns.Count - 1; j++)
                {
                    dgv_vrem.Rows[0].Cells[j].Value = Math.Round(Convert.ToDecimal(dgv_vrem.Rows[0].Cells[j].Value) + Convert.ToDecimal(dgv_vrem.Rows[i].Cells[j].Value), 3);
                    //MessageBox.Show(dgv_vrem.Rows[0].Cells[j].Value.ToString());
                }
            }
            //return dgv;
            // return dgv_vrem;
        } //суммирование данных для вывода суммы в конце таблицы
        public DataGridView nebalans_sutki_finish(DataGridView dgv, DataGridView dgv_vrem, CheckBox cb_niz, ComboBox cb_rp, DateTimePicker dtp, Chart ch, ContextMenuStrip cms)
        {
            cms.Items[4].Enabled = false;
            cms.Items[6].Enabled = false;
            cms.Items[7].Enabled = true;

            if (cb_niz.Checked == false) { ru = " РУ - 6-10 кВ"; }
            if (cb_niz.Checked == true) { ru = " РУ - 0,4 кВ"; }

            trans.Clear();
            List<class10> tr = new List<class10>();

            for (int i = 0; i < 1; i++)
            {
                for (int j = 1; j < dgv.Columns.Count - 1; j++)
                {


                    class10 m = new class10();
                    //m.rp = "";
                    // m.fider = dataGridView3.Rows[i].Cells[0].Value.ToString();
                    //m.vvod = dataGridView1.Rows[i].Cells[0].Value.ToString();
                    m.data = dgv.Columns[j].HeaderText.ToString();
                    m.val_f = dgv.Rows[i].Cells[j].Value.ToString();
                    m.val_v = dgv_vrem.Rows[i].Cells[j].Value.ToString();
                    m.raznica = (Convert.ToDecimal(dgv_vrem.Rows[i].Cells[j].Value) - Convert.ToDecimal(dgv.Rows[i].Cells[j].Value)).ToString();
                    trans.Add(m);


                }
            }

            dgv.DataSource = trans;
            for (int k = 0; k < dgv.Rows.Count; k++)
            {
                dgv.Rows[k].Cells[0].Value = cb_rp.SelectedItem.ToString();
                try
                {
                    dgv.Rows[k].Cells[5].Value = (Math.Round((Convert.ToDecimal(dgv.Rows[k].Cells[4].Value) / Convert.ToDecimal(dgv.Rows[k].Cells[3].Value)) * 100, 3)).ToString();
                }
                catch { dgv.Rows[k].Cells[5].Value = "0"; }
            }
            //dgv.RowHeadersWidth = 10;

            dgv.Columns[1].HeaderText = "Время";
            dgv.Columns[0].HeaderText = "Имя";
            dgv.Columns[2].HeaderText = "Отход(кВт*ч)";
            dgv.Columns[3].HeaderText = "Ввода(кВт*ч)";
            dgv.Columns[4].HeaderText = "Баланс(кВт*ч)";
            dgv.Columns[5].HeaderText = "Баланс(%)";
            dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            //sum(sender, e, dgv, 4, label1);
            // sum(sender, e, dataGridView3, 5, label6);

            int ss = 0;
            ch.Series[0].Points.Clear();
            while (ss < dgv.Rows.Count)
            {
                var skip = dgv.Rows[ss].Cells[4].Value;
                var date = dgv.Rows[ss].Cells[1].Value;
                //string s = Convert.ToString(date);
                //string s1 = s.Substring(s.IndexOf(" "),8);
                ch.ChartAreas[0].AxisY.MajorGrid.Interval = 5;
                ch.ChartAreas[0].AxisX.LabelStyle.Format = "HH:mm";
                ch.ChartAreas[0].AxisX.IntervalType = System.Windows.Forms.DataVisualization.Charting.DateTimeIntervalType.Hours;
                ch.ChartAreas[0].AxisX.Interval = 1;
                ch.Titles[0].Text = "Информация о суточном небалансе " + cb_rp.SelectedItem.ToString() + "  дата  " + dtp.Value.ToString("yyyy-MM-dd") + ru;
                foreach (Legend objLegend in ch.Legends) //убрали легенду из графика
                {
                    objLegend.Enabled = false;
                }

                ch.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                ch.Series[0].Points.AddXY(Convert.ToDateTime(date), Convert.ToDecimal(skip));
                ch.Series[0].ToolTip = "X=#VALX, Y=#VALY";
                ss++;
            }
            return dgv;
        } //небаланс за сутки по РП
        public DataGridView nebalans_sutki_finish_prisoed(DataGridView dgv_f, DataGridView dgv_v)
        {
            List<class10> tr = new List<class10>();
            trans.Clear();

            //фидера

            for (int i = 0; i < dgv_f.Rows.Count; i++)
            {
                for (int j = 1; j < dgv_f.Columns.Count - 1; j++)
                {
                    class10 m = new class10();

                    m.rp = dgv_f.Rows[i].Cells[0].Value.ToString();
                    if (dgv_f.Columns[j].HeaderText.Substring(0, 1).Equals("0"))
                    {
                        m.data = dgv_f.Columns[j].HeaderText.Remove(0, 1).ToString();
                    }
                    else { m.data = dgv_f.Columns[j].HeaderText.ToString(); }
                    m.val_f = dgv_f.Rows[i].Cells[j].Value.ToString();
                    trans.Add(m);
                }
            }

            //ввода

            for (int i = 0; i < dgv_v.Rows.Count; i++)
            {
                for (int j = 1; j < dgv_v.Columns.Count - 1; j++)
                {
                    class10 m = new class10();

                    m.rp = dgv_v.Rows[i].Cells[0].Value.ToString();
                    if (dgv_v.Columns[j].HeaderText.Substring(0, 1).Equals("0"))
                    {
                        m.data = dgv_v.Columns[j].HeaderText.Remove(0, 1).ToString();

                    }
                    else { m.data = dgv_v.Columns[j].HeaderText.ToString(); }
                    m.val_f = dgv_v.Rows[i].Cells[j].Value.ToString();
                    trans.Add(m);
                }
            }



            dgv_f.DataSource = trans;
            dgv_f.Columns[0].HeaderText = "Имя";
            dgv_f.Columns[1].HeaderText = "Время";
            dgv_f.Columns[2].HeaderText = "Энергия(кВт*ч)";
            dgv_f.Columns[3].Visible = false;
            dgv_f.Columns[4].Visible = false;
            dgv_f.Columns[5].Visible = false;
            dgv_f.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
            dgv_f.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dgv_f.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            return dgv_f;
        }   //небаланс при выборе присоединений
        public DataGridView nebalans_sutki_finish_prisoed_povremeni(DataGridView dgv, ComboBox cb)
        {
            gr.Clear();
            string s;
            for (int i = 0; i < dgv.RowCount; i++)
            {
                //dataGridView3.Rows[i].Visible = true;
                s = dgv.Rows[i].Cells[1].Value.ToString();
                if (s.Equals(cb.SelectedItem.ToString()))
                //if (s.Substring(s.IndexOf(" ") + 1, s.Length - s.IndexOf(" ") - 4).Equals(comboBox2.SelectedItem.ToString()))
                {

                    chas m = new chas();
                    m.name_f = dgv.Rows[i].Cells[0].Value.ToString();
                    m.t_time = dgv.Rows[i].Cells[1].Value.ToString();
                    m.v_value = dgv.Rows[i].Cells[2].Value.ToString();
                    gr.Add(m);
                }
                //else { dataGridView3.Rows[i].Visible = false; }
            }
            dgv.DataSource = null;
            dgv.DataSource = gr;
            dgv.Columns[1].HeaderText = "Время";
            dgv.Columns[0].HeaderText = "Имя";
            dgv.Columns[2].HeaderText = "Энергия(кВт*ч)";
            dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
            //dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            //dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            return dgv;

        }  //небаланс за сутки по времени
        public DataGridView nebalans_sutki_finish_prisoed_poothod(DataGridView dgv, ComboBox cb_othod, Chart ch, CheckBox cbox, DateTimePicker dtp, ContextMenuStrip cms)
        {
            of.Clear();
            cms.Items[4].Enabled = false;
            cms.Items[6].Enabled = false;
            cms.Items[7].Enabled = true;
            string s;
            for (int i = 0; i < dgv.RowCount; i++)
            {
                //dataGridView3.Rows[i].Visible = true;
                s = dgv.Rows[i].Cells[0].Value.ToString();

                if (s.Equals(cb_othod.SelectedItem.ToString()))

                //if (s.Substring(s.IndexOf(" ") + 1, s.Length - s.IndexOf(" ") - 4).Equals(comboBox2.SelectedItem.ToString()))
                {

                    one_fider m = new one_fider();
                    m.name_f = dgv.Rows[i].Cells[0].Value.ToString();
                    m.t_time = dgv.Rows[i].Cells[1].Value.ToString();
                    m.v_value = dgv.Rows[i].Cells[2].Value.ToString();
                    of.Add(m);
                }

                //else { dataGridView3.Rows[i].Visible = false; }
            }
            var sum = of.Sum(p => Convert.ToDecimal(p.v_value));
            one_fider o = new one_fider();
            o.name_f = "Итог";
            o.t_time = "Сутки";
            o.v_value = sum.ToString();
            of.Add(o);

            dgv.DataSource = null;
            dgv.DataSource = of;
            dgv.Columns[1].HeaderText = "Время";
            dgv.Columns[0].HeaderText = "Имя";
            dgv.Columns[2].HeaderText = "Энергия(кВт*ч)";
            dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True;



            int ss = 0;
            ch.Series[0].Points.Clear();
            while (ss < dgv.Rows.Count - 1)
            {
                var skip = dgv.Rows[ss].Cells[2].Value;
                var date = dgv.Rows[ss].Cells[1].Value;

                //string s = Convert.ToString(date);
                //string s1 = s.Substring(s.IndexOf(" "),8);
                //chart1.ChartAreas[0].AxisY.MajorGrid.Interval = 5;
                ch.ChartAreas[0].AxisX.LabelStyle.Format = "HH:mm";
                ch.ChartAreas[0].AxisX.IntervalType = System.Windows.Forms.DataVisualization.Charting.DateTimeIntervalType.Hours;
                ch.ChartAreas[0].AxisX.Interval = 1;
                if (cbox.Checked == false) { ru = " РУ - 6-10 кВ"; }
                if (cbox.Checked == true) { ru = " РУ - 0,4 кВ"; }
                ch.Titles[0].Text = "Информация о суточном потреблении " + cb_othod.SelectedItem.ToString() + "  дата  " + dtp.Value.ToString("yyyy-MM-dd") + ru;
                foreach (Legend objLegend in ch.Legends) //убрали легенду из графика
                {
                    objLegend.Enabled = false;
                }

                ch.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                ch.Series[0].Points.AddXY(Convert.ToDateTime(date), Convert.ToDecimal(skip));
                ch.Series[0].ToolTip = "X=#VALX, Y=#VALY";
                ss++;
            }
            return dgv;
        }  //небаланс за сутки по присоединениям
        public void print_page(System.Drawing.Printing.PrintPageEventArgs sdpa, DataGridView dgv, ComboBox cb, string zagolovok)
        {
            try
            {
                //Set the left margin
                int iLeftMargin = sdpa.MarginBounds.Left;
                //Set the top margin
                int iTopMargin = sdpa.MarginBounds.Top;
                //Whether more pages have to print or not
                bool bMorePagesToPrint = false;
                int iTmpWidth = 0;

                //For the first page to print set the cell width and header height
                if (bFirstPage)
                {
                    foreach (DataGridViewColumn GridCol in dgv.Columns)
                    {
                        iTmpWidth = (int)(Math.Floor((double)((double)GridCol.Width /
                                       (double)iTotalWidth * (double)iTotalWidth *
                                       ((double)sdpa.MarginBounds.Width / (double)iTotalWidth))));

                        iHeaderHeight = (int)(sdpa.Graphics.MeasureString(GridCol.HeaderText,
                                    GridCol.InheritedStyle.Font, iTmpWidth).Height) + 11;

                        // Save width and height of headres
                        arrColumnLefts.Add(iLeftMargin);
                        arrColumnWidths.Add(iTmpWidth);
                        iLeftMargin += iTmpWidth;
                    }
                }
                //Loop till all the grid rows not get printed
                while (iRow <= dgv.Rows.Count - 1)
                {
                    DataGridViewRow GridRow = dgv.Rows[iRow];
                    //Set the cell height
                    iCellHeight = GridRow.Height - 3;
                    int iCount = 0;
                    //Check whether the current page settings allo more rows to print
                    if (iTopMargin + iCellHeight >= sdpa.MarginBounds.Height + sdpa.MarginBounds.Top)
                    {
                        bNewPage = true;
                        bFirstPage = false;
                        bMorePagesToPrint = true;
                        break;
                    }
                    else
                    {
                        if (bNewPage)
                        {
                            //Draw Header
                            sdpa.Graphics.DrawString(zagolovok, new Font(dgv.Font, FontStyle.Bold),
                                    Brushes.Black, sdpa.MarginBounds.Left, sdpa.MarginBounds.Top -
                                    sdpa.Graphics.MeasureString(zagolovok, new Font(dgv.Font,
                                    FontStyle.Bold), sdpa.MarginBounds.Width).Height - 13);

                            String strDate = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToShortTimeString();
                            //Draw Date
                            //e.Graphics.DrawString(strDate, new Font(dataGridView3.Font, FontStyle.Bold),
                            //        Brushes.Black, e.MarginBounds.Left + (e.MarginBounds.Width -
                            //        e.Graphics.MeasureString(strDate, new Font(dataGridView3.Font,
                            //        FontStyle.Bold), e.MarginBounds.Width).Width), e.MarginBounds.Top -
                            //        e.Graphics.MeasureString("Customer Summary", new Font(new Font(dataGridView3.Font,
                            //        FontStyle.Bold), FontStyle.Bold), e.MarginBounds.Width).Height - 13);

                            //Draw Columns                 
                            iTopMargin = sdpa.MarginBounds.Top;
                            foreach (DataGridViewColumn GridCol in dgv.Columns)
                            {
                                sdpa.Graphics.FillRectangle(new SolidBrush(Color.LightGray),
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                sdpa.Graphics.DrawRectangle(Pens.Black,
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                sdpa.Graphics.DrawString(GridCol.HeaderText, GridCol.InheritedStyle.Font,
                                    new SolidBrush(GridCol.InheritedStyle.ForeColor),
                                    new RectangleF((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight), strFormat);
                                iCount++;
                            }
                            bNewPage = false;
                            iTopMargin += iHeaderHeight;
                        }
                        iCount = 0;
                        //Draw Columns Contents                
                        foreach (DataGridViewCell Cel in GridRow.Cells)
                        {
                            if (Cel.Value != null)
                            {
                                sdpa.Graphics.DrawString(Cel.Value.ToString(), Cel.InheritedStyle.Font,
                                            new SolidBrush(Cel.InheritedStyle.ForeColor),
                                            new RectangleF((int)arrColumnLefts[iCount], (float)iTopMargin,
                                            (int)arrColumnWidths[iCount], (float)iCellHeight), strFormat);
                            }
                            //Drawing Cells Borders 
                            sdpa.Graphics.DrawRectangle(Pens.Black, new Rectangle((int)arrColumnLefts[iCount],
                                    iTopMargin, (int)arrColumnWidths[iCount], iCellHeight));

                            iCount++;
                        }
                    }
                    iRow++;
                    iTopMargin += iCellHeight;
                }

                //If more lines exist, print another page.
                if (bMorePagesToPrint)
                    sdpa.HasMorePages = true;
                else
                    sdpa.HasMorePages = false;
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        } //для печати
        public void print_begin(DataGridView dgv)
        {
            try
            {
                strFormat = new StringFormat();
                strFormat.Alignment = StringAlignment.Near;
                strFormat.LineAlignment = StringAlignment.Center;
                strFormat.Trimming = StringTrimming.EllipsisCharacter;

                arrColumnLefts.Clear();
                arrColumnWidths.Clear();
                iCellHeight = 0;
                iRow = 0;
                bFirstPage = true;
                bNewPage = true;

                // Calculating Total Widths
                iTotalWidth = 0;
                foreach (DataGridViewColumn dgvGridCol in dgv.Columns)
                {
                    iTotalWidth += dgvGridCol.Width;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        } //для печати
        public void vid_graphic(ComboBox cb, Chart ch1, Chart ch2, Chart ch3)
        {
            if (cb.Text == "bar")
            {
                ch1.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Bar;
                ch2.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Bar;
                ch3.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Bar;
            }
            if (cb.Text == "spline")
            {
                ch1.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
                ch2.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
                ch3.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            }
            if (cb.Text == "line")
            {
                ch1.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                ch2.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                ch3.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            }
            if (cb.Text == "column")
            {
                ch1.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                ch2.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                ch3.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
            }
        } //вид графика
        public void vid_graphic_proc_nebalans(Chart ch, ComboBox cb)
        {
            int ss = 0;
            ch.Series[0].Points.Clear();

            while (ss < b.Count)
            {

                var proc = b1[ss];
                var skip = b[ss];
                var date = dat[ss];
                //chart1.ChartAreas[0].AxisY.MajorGrid.Interval = 50;
                //chart1.ChartAreas[0].AxisY.Title = "кВт*ч";
                ch.ChartAreas[0].AxisX.LabelStyle.Format = "dd";
                ch.ChartAreas[0].AxisX.IntervalType = System.Windows.Forms.DataVisualization.Charting.DateTimeIntervalType.Days;
                ch.ChartAreas[0].AxisX.Interval = 1;
                foreach (Legend objLegend in ch.Legends) //убрали легенду из графика
                {
                    objLegend.Enabled = false;
                }

                if (cb.Text == "%")
                {

                    ch.Series[0].Points.AddXY(Convert.ToDateTime(date), Convert.ToDouble(proc));
                    ch.Series[0].ToolTip = "X=#VALX, Y=#VALY";
                    ch.ChartAreas[0].AxisY.Title = "%";


                }
                if (cb.Text == "кВт*ч")
                {
                    ch.Series[0].Points.AddXY(Convert.ToDateTime(date), Convert.ToDouble(skip));
                    ch.Series[0].ToolTip = "X=#VALX, Y=#VALY";
                    ch.ChartAreas[0].AxisY.Title = "кВт*ч";
                }
                ss++;
            }
        } //график в процентах небаланс за период
        public void vid_graphic_proc_sut(Chart ch, DataGridView dgv, ComboBox cb)
        {

            int ss = 0;
            ch.Series[0].Points.Clear();

            while (ss < dgv.Rows.Count)
            {

                var proc = dgv.Rows[ss].Cells[5].Value;
                var skip = dgv.Rows[ss].Cells[4].Value;
                var date = dgv.Rows[ss].Cells[1].Value;
                //chart1.ChartAreas[0].AxisY.MajorGrid.Interval = 5;
                ch.ChartAreas[0].AxisX.LabelStyle.Format = "HH:mm";
                ch.ChartAreas[0].AxisX.IntervalType = System.Windows.Forms.DataVisualization.Charting.DateTimeIntervalType.Hours;
                ch.ChartAreas[0].AxisX.Interval = 1;
                foreach (Legend objLegend in ch.Legends) //убрали легенду из графика
                {
                    objLegend.Enabled = false;
                }

                if (cb.Text == "%")
                {
                    ch.ChartAreas[0].AxisY.MajorGrid.Interval = 0.5;
                    ch.Series[0].Points.AddXY(Convert.ToDateTime(date), Convert.ToDecimal(proc));
                    ch.Series[0].ToolTip = "X=#VALX, Y=#VALY";
                    ch.ChartAreas[0].AxisY.Title = "%";


                }
                if (cb.Text == "кВт*ч")
                {
                    ch.ChartAreas[0].AxisY.MajorGrid.Interval = 5;
                    ch.Series[0].Points.AddXY(Convert.ToDateTime(date), Convert.ToDecimal(skip));
                    ch.Series[0].ToolTip = "X=#VALX, Y=#VALY";
                    ch.ChartAreas[0].AxisY.Title = "кВт*ч";
                }
                ss++;
            }
        } //график в процентах небаланс за сутки
        public DataGridView povernut_prisoed(DataGridView dgv)
        {
            asc_p.Clear();
            List<String> ls1 = new List<String>();
            for (int i = 0; i < trans.Count - 1; i++)
            {
                if (ls1.IndexOf(trans[i].rp) < 0)
                {
                    ls1.Add(trans[i].rp);
                    povernut m = new povernut();
                    m.name_rp = trans[i].rp;
                    asc_p.Add(m);
                }
            }

            for (int k = 0; k < trans.Count; k++)
            {
                for (int j = 0; j < asc_p.Count; j++)
                {
                    if (trans[k].rp.Equals(asc_p[j].name_rp))
                    {
                        string streng = trans[k].data.ToString().Trim();
                        asc_p[j].v(timeparser_pris(streng), trans[k].val_f);

                    }
                }
            }

            dgv.DataSource = null;
            dgv.DataSource = asc_p;
            dgv.Columns[0].HeaderText = "Название";  //присваиваем текст к заголовку колонок
            dgv.Columns[1].HeaderText = "00:00";
            dgv.Columns[2].HeaderText = "00:30";
            dgv.Columns[3].HeaderText = "01:00";
            dgv.Columns[4].HeaderText = "01:30";
            dgv.Columns[5].HeaderText = "02:00";
            dgv.Columns[6].HeaderText = "02:30";
            dgv.Columns[7].HeaderText = "03:00";
            dgv.Columns[8].HeaderText = "03:30";
            dgv.Columns[9].HeaderText = "04:00";
            dgv.Columns[10].HeaderText = "04:30";
            dgv.Columns[11].HeaderText = "05:00";
            dgv.Columns[12].HeaderText = "05:30";
            dgv.Columns[13].HeaderText = "06:00";
            dgv.Columns[14].HeaderText = "06:30";
            dgv.Columns[15].HeaderText = "07:00";
            dgv.Columns[16].HeaderText = "07:30";
            dgv.Columns[17].HeaderText = "08:00";
            dgv.Columns[18].HeaderText = "08:30";
            dgv.Columns[19].HeaderText = "09:00";
            dgv.Columns[20].HeaderText = "09:30";
            dgv.Columns[21].HeaderText = "10:00";
            dgv.Columns[22].HeaderText = "10:30";
            dgv.Columns[23].HeaderText = "11:00";
            dgv.Columns[24].HeaderText = "11:30";
            dgv.Columns[25].HeaderText = "12:00";
            dgv.Columns[26].HeaderText = "12:30";
            dgv.Columns[27].HeaderText = "13:00";
            dgv.Columns[28].HeaderText = "13:30";
            dgv.Columns[29].HeaderText = "14:00";
            dgv.Columns[30].HeaderText = "14:30";
            dgv.Columns[31].HeaderText = "15:00";
            dgv.Columns[32].HeaderText = "15:30";
            dgv.Columns[33].HeaderText = "16:00";
            dgv.Columns[34].HeaderText = "16:30";
            dgv.Columns[35].HeaderText = "17:00";
            dgv.Columns[36].HeaderText = "17:30";
            dgv.Columns[37].HeaderText = "18:00";
            dgv.Columns[38].HeaderText = "18:30";
            dgv.Columns[39].HeaderText = "19:00";
            dgv.Columns[40].HeaderText = "19:30";
            dgv.Columns[41].HeaderText = "20:00";
            dgv.Columns[42].HeaderText = "20:30";
            dgv.Columns[43].HeaderText = "21:00";
            dgv.Columns[44].HeaderText = "21:30";
            dgv.Columns[45].HeaderText = "22:00";
            dgv.Columns[46].HeaderText = "22:30";
            dgv.Columns[47].HeaderText = "23:00";
            dgv.Columns[48].HeaderText = "23:30";
            dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
            return dgv;

        } //изменить вид присоед
        public void nebalans_sutki_energomera(DataGridView dgv, DataGridView dgv_vrem, CheckBox cb_niz, ComboBox cb_rp, DateTimePicker dtp, RichTextBox rtb)
        {
            prov = 1;
            dgv.DataSource = null;
            dgv_vrem.DataSource = null;
            dgv.Rows.Clear();
            dgv_vrem.Rows.Clear();
            list1.Clear();

            string vv = "";
            string ff = "";

            if (cb_niz.Checked == true) { vv = "vvod_niz"; ff = "fider_niz"; }
            if (cb_niz.Checked == false) { vv = "vvod"; ff = "fider"; }

            DateTime d1 = new DateTime(dtp.Value.Year, dtp.Value.Month, dtp.Value.Day); //дата начала периода
            double x1 = d1.ToOADate();    //перевод нормальной даты в абсолютную
            string chas1 = "0:00";
            chas1 = chas1.Substring(0, chas1.IndexOf(":"));
            string min1 = "0:00";
            min1 = min1.Substring(min1.IndexOf(":") + 1, 2);
            double ab_vr1 = (Convert.ToDouble(chas1) * 60 + Convert.ToDouble(min1)) * 0.000694443; //перевод часов и минут (то есть время) в абсолютное. Это от
            string chas2 = "23:30";
            chas2 = chas2.Substring(0, chas2.IndexOf(":"));
            string min2 = "23:30";
            min2 = min2.Substring(min2.IndexOf(":") + 1, 2);
            double ab_vr2 = (Convert.ToDouble(chas2) * 60 + 2 * Convert.ToDouble(min2)) * 0.000694445;  //перевод часов и минут (то есть время) в абсолютное. Это до
            double vr1 = x1 + ab_vr1;
            double vr3 = x1 + ab_vr2;

            string vr11 = Convert.ToString(vr1).Replace(",", ".");
            string vr22 = Convert.ToString(vr3).Replace(",", ".");

            string n_d = DateTime.FromOADate(vr1).ToString().Replace(".", "-");
            string k_d = DateTime.FromOADate(vr3).ToString().Replace(".", "-");

            NpgsqlConnection connection1 = new NpgsqlConnection(monitor_options);
            connection1.Open();
            NpgsqlConnection conn = new NpgsqlConnection(energomera_comobjects);
            DataTable table11 = new DataTable();
            conn.Open();
            NpgsqlCommand com11 = new NpgsqlCommand(" select  balancegroup.\"Name\", valuesmeters.\"SerialNumber\", valuesmeters.\"MeterId\" from (SELECT  \"Meters\".\"SerialNumber\", \"Meters\".\"MeterId\" from \"Meters\" where \"Meters\".\"SerialNumber\" IN (select \"SerialNumber\" from \"Meters\" where \"ComObjectId\" IN (select \"ComObjectId\" from \"ComObjects\" where \"Name\" = '" + cb_rp.SelectedItem.ToString() + "'))) as valuesmeters left join (SELECT  \"BalanceGroups_Meters\".\"BalanceGroupId\", \"BalanceGroups_Meters\".\"MeterId\", \"BalanceGroups\".\"Name\" FROM \"BalanceGroups_Meters\", \"BalanceGroups\" WHERE \"BalanceGroups\".\"BalanceGroupId\"=\"BalanceGroups_Meters\".\"BalanceGroupId\") as balancegroup on balancegroup.\"MeterId\"=valuesmeters.\"MeterId\" ", conn);
            rtb.Text = rtb.Text + " select  balancegroup.\"Name\", valuesmeters.\"SerialNumber\", valuesmeters.\"MeterId\" from (SELECT  \"Meters\".\"SerialNumber\", \"Meters\".\"MeterId\" from \"Meters\" where \"Meters\".\"SerialNumber\" IN (select \"SerialNumber\" from \"Meters\" where \"ComObjectId\" IN (select \"ComObjectId\" from \"ComObjects\" where \"Name\" = '" + cb_rp.SelectedItem.ToString() + "'))) as valuesmeters left join (SELECT  \"BalanceGroups_Meters\".\"BalanceGroupId\", \"BalanceGroups_Meters\".\"MeterId\", \"BalanceGroups\".\"Name\" FROM \"BalanceGroups_Meters\", \"BalanceGroups\" WHERE \"BalanceGroups\".\"BalanceGroupId\"=\"BalanceGroups_Meters\".\"BalanceGroupId\") as balancegroup on balancegroup.\"MeterId\"=valuesmeters.\"MeterId\" \n \n";
            NpgsqlDataAdapter dap11 = new NpgsqlDataAdapter(com11);
            dap11.Fill(table11);
            list1_energo.Clear();

            foreach (DataRow row in table11.Rows)
            {
                class2_energo m = new class2_energo();
                m.meters_id = row.ItemArray[2].ToString();
                m.name = cb_rp.SelectedItem.ToString();
                m.name_f = row.ItemArray[0].ToString();
                m.sn = row.ItemArray[1].ToString();
                list1_energo.Add(m);
            }

            string rp = cb_rp.SelectedItem.ToString();
            List<string> list_f = new List<string>(); //лист фидеров
            List<string> list_v = new List<string>(); //лист вводов

            NpgsqlCommand com1 = new NpgsqlCommand("select \"name_tii\", \"ktt\"  from balans where \"name_rp\" = '" + rp + "' and \"" + ff + "\" = TRUE", connection1);
            rtb.Text = rtb.Text + "select \"name_tii\", \"ktt\"  from balans where \"name_rp\" = '" + rp + "' and \"" + ff + "\" = TRUE \n \n";
            NpgsqlDataReader datar1 = com1.ExecuteReader();

            fider_koef.Clear();

            while (datar1.Read())
            {
                class8 m = new class8();
                m.name_f = datar1.GetValue(0).ToString();
                m.koef = datar1.GetValue(1).ToString();
                fider_koef.Add(m);
                list_f.Add(datar1.GetValue(0).ToString());
            }

            vvod_koef.Clear();
            com1 = new NpgsqlCommand("select \"name_tii\", \"ktt\"  from balans where \"name_rp\" = '" + rp + "' and \"" + vv + "\" = TRUE", connection1);
            rtb.Text = rtb.Text + "select \"name_tii\", \"ktt\"  from balans where \"name_rp\" = '" + rp + "' and \"" + vv + "\" = TRUE \n \n";
            datar1 = com1.ExecuteReader();
            vvod_koef.Clear();
            while (datar1.Read())
            {
                class8_1 m = new class8_1();
                m.name_f = datar1.GetValue(0).ToString();
                m.koef = datar1.GetValue(1).ToString();
                vvod_koef.Add(m);
                list_v.Add(datar1.GetValue(0).ToString());
            }

            for (int i = 0; i < fider_koef.Count; i++)
            {

                NpgsqlCommand com4 = new NpgsqlCommand("select \"MeterId\" from \"BalanceGroups_Meters\" where \"BalanceGroupId\" IN (select \"BalanceGroupId\" from \"BalanceGroups\" where \"Name\" IN ('" + list_f[i] + "'))", conn);
                rtb.Text = rtb.Text + "select \"MeterId\" from \"BalanceGroups_Meters\" where \"BalanceGroupId\" IN (select \"BalanceGroupId\" from \"BalanceGroups\" where \"Name\" IN ('" + list_f[i] + "')) \n \n";
                list_f[i] = Convert.ToString(com4.ExecuteScalar());
            }

            for (int i = 0; i < vvod_koef.Count; i++)
            {

                NpgsqlCommand com3 = new NpgsqlCommand("select \"MeterId\" from \"BalanceGroups_Meters\" where \"BalanceGroupId\" IN (select \"BalanceGroupId\" from \"BalanceGroups\" where \"Name\" IN ('" + list_v[i] + "'))", conn);
                rtb.Text = rtb.Text + "select \"MeterId\" from \"BalanceGroups_Meters\" where \"BalanceGroupId\" IN (select \"BalanceGroupId\" from \"BalanceGroups\" where \"Name\" IN ('" + list_v[i] + "')) \n \n";
                list_v[i] = Convert.ToString(com3.ExecuteScalar());

            }

            connection1.Close();
            conn.Close();

            StringBuilder v_id = new StringBuilder();
            v_id.Append("(");
            for (int i = 0; i < list_f.Count; i++)
            {
                v_id.Append(list_f[i]);

                if (i != list_f.Count - 1)
                {
                    v_id.Append(", ");
                }
            }
            v_id.Append(")");

            StringBuilder v1_id = new StringBuilder();
            v1_id.Append("(");
            for (int i = 0; i < list_v.Count; i++)
            {
                v1_id.Append(list_v[i]);

                if (i != list_v.Count - 1)
                {
                    v1_id.Append(", ");
                }
            }
            v1_id.Append(")");

            NpgsqlConnection conn1 = new NpgsqlConnection(energomera_comobjects);
            DataTable table = new DataTable();
            conn1.Open();
            NpgsqlCommand com = new NpgsqlCommand(" SELECT  \"Meters\".\"SerialNumber\", \"Meters\".\"MeterId\",  \"ValueProfiles\".\"DT\", \"ValueProfiles\".\"Val\" from  \"ValueProfiles\", \"Meters\" where \"ValueProfiles\".\"MeterId\" IN  " + v_id + " and \"ValueProfiles\".\"MeterId\"=\"Meters\".\"MeterId\" and \"ValueProfiles\".\"StateId\" = 1  and \"ValueProfiles\".\"DT\" >= '" + n_d + "' and \"ValueProfiles\".\"DT\" <='" + k_d + "' order by \"ValueProfiles\".\"MeterId\", \"ValueProfiles\".\"DT\" ", conn1);
            //rtb.Text = rtb.Text + " SELECT  \"Meters\".\"SerialNumber\", \"Meters\".\"MeterId\",  \"Values\".\"DT\", \"Values\".\"Val\" from  \"Values\", \"Meters\" where \"Values\".\"MeterId\" IN  " + v_id + " and \"Values\".\"MeterId\"=\"Meters\".\"MeterId\" and \"Values\".\"TariffId\" = 1  and \"Values\".\"DT\" >= '" + n_d + "' and \"Values\".\"DT\" < '" + k_d + "' order by \"Values\".\"MeterId\", \"Values\".\"DT\" \n \n";
            rtb.Text = rtb.Text + " SELECT  \"Meters\".\"SerialNumber\", \"Meters\".\"MeterId\",  \"ValueProfiles\".\"DT\", \"ValueProfiles\".\"Val\" from  \"ValueProfiles\", \"Meters\" where \"ValueProfiles\".\"MeterId\" IN  " + v_id + " and \"ValueProfiles\".\"MeterId\"=\"Meters\".\"MeterId\" and \"ValueProfiles\".\"StateId\" = 1  and \"ValueProfiles\".\"DT\" >= '" + n_d + "' and \"ValueProfiles\".\"DT\" <= '" + k_d + "' order by \"ValueProfiles\".\"MeterId\", \"ValueProfiles\".\"DT\" ";
            NpgsqlDataAdapter dap = new NpgsqlDataAdapter(com);
            dap.Fill(table);
            list_energo.Clear();
            foreach (DataRow row in table.Rows)
            {
                class0_energo m = new class0_energo();
                m.t_time = row.ItemArray[2].ToString();
                m.meters_id = row.ItemArray[1].ToString();
                m.v_value = row.ItemArray[3].ToString();
                m.sn = row.ItemArray[0].ToString();
                list_energo.Add(m);
            }
            conn1.Close();

            NpgsqlConnection conn2 = new NpgsqlConnection(energomera_comobjects);
            DataTable table2 = new DataTable();
            conn2.Open();
            NpgsqlCommand com2 = new NpgsqlCommand(" SELECT  \"Meters\".\"SerialNumber\", \"Meters\".\"MeterId\",  \"ValueProfiles\".\"DT\", \"ValueProfiles\".\"Val\" from  \"ValueProfiles\", \"Meters\" where \"ValueProfiles\".\"MeterId\" IN  " + v1_id + " and \"ValueProfiles\".\"MeterId\"=\"Meters\".\"MeterId\" and \"ValueProfiles\".\"StateId\" = 1  and \"ValueProfiles\".\"DT\" >= '" + n_d + "' and \"ValueProfiles\".\"DT\" <= '" + k_d + "' order by \"ValueProfiles\".\"MeterId\", \"ValueProfiles\".\"DT\" ", conn2);
            //rtb.Text = rtb.Text + " SELECT  \"Meters\".\"SerialNumber\", \"Meters\".\"MeterId\",  \"Values\".\"DT\", \"Values\".\"Val\" from  \"Values\", \"Meters\" where \"Values\".\"MeterId\" IN  " + v1_id + " and \"Values\".\"MeterId\"=\"Meters\".\"MeterId\" and \"Values\".\"TariffId\" = 1  and \"Values\".\"DT\" >= '" + n_d + "' and \"Values\".\"DT\" < '" + k_d + "' order by \"Values\".\"MeterId\", \"Values\".\"DT\" \n \n";
            rtb.Text = rtb.Text + " SELECT  \"Meters\".\"SerialNumber\", \"Meters\".\"MeterId\",  \"ValueProfiles\".\"DT\", \"ValueProfiles\".\"Val\" from  \"ValueProfiles\", \"Meters\" where \"ValueProfiles\".\"MeterId\" IN  " + v1_id + " and \"ValueProfiles\".\"MeterId\"=\"Meters\".\"MeterId\" and \"ValueProfiles\".\"StateId\" = 1  and \"ValueProfiles\".\"DT\" >= '" + n_d + "' and \"ValueProfiles\".\"DT\" <='" + k_d + "' order by \"ValueProfiles\".\"MeterId\", \"ValueProfiles\".\"DT\" ";
            NpgsqlDataAdapter dap2 = new NpgsqlDataAdapter(com2);
            dap2.Fill(table2);
            list_vlist_energo.Clear();
            foreach (DataRow row in table2.Rows)
            {
                class9_energo m = new class9_energo();
                m.t_time = row.ItemArray[2].ToString();
                m.meters_id = row.ItemArray[1].ToString();
                m.v_value = row.ItemArray[3].ToString();
                m.sn = row.ItemArray[0].ToString();
                list_vlist_energo.Add(m);
            }
            conn2.Close();

            vivod_energomera.Clear();
            vivod_energomera_v.Clear();

            for (int i = 0; i < list_energo.Count; i++)
            {
                for (int j = 0; j < list1_energo.Count; j++)
                {
                    if (list_energo[i].meters_id == list1_energo[j].meters_id)
                    {
                        energomera m = new energomera();
                        m.name_rp = list1_energo[j].name;
                        m.name_f = list1_energo[j].name_f;
                        m.name_sch = list1_energo[j].sn;
                        m.t_time = list_energo[i].t_time;
                        m.v_value = list_energo[i].v_value;
                        vivod_energomera.Add(m);
                    }
                }
            }

            for (int i = 0; i < list_vlist_energo.Count; i++)
            {
                for (int j = 0; j < list1_energo.Count; j++)
                {
                    if (list_vlist_energo[i].meters_id == list1_energo[j].meters_id)
                    {
                        energomera_v m = new energomera_v();
                        m.name_rp = list1_energo[j].name;
                        m.name_f = list1_energo[j].name_f;
                        m.name_sch = list1_energo[j].sn;
                        m.t_time = list_vlist_energo[i].t_time;
                        m.v_value = list_vlist_energo[i].v_value;

                        vivod_energomera_v.Add(m);
                    }
                }
            }

            asc.Clear();
            asc_v.Clear();
            List<String> ls1 = new List<String>();
            List<String> ls2 = new List<String>();

            for (int i = 0; i < vivod_energomera.Count - 1; i++)
            {
                if (ls1.IndexOf(vivod_energomera[i].name_f) < 0)
                {
                    ls1.Add(vivod_energomera[i].name_f);
                    class4 m = new class4();
                    m.name_rp = vivod_energomera[i].name_f;
                    asc.Add(m);
                }
            }

            for (int k = 0; k < vivod_energomera.Count; k++)
            {

                for (int j = 0; j < asc.Count; j++)
                {

                    if (vivod_energomera[k].name_f.Equals(asc[j].name_rp))
                    {

                        string streng = vivod_energomera[k].t_time.Substring(vivod_energomera[k].t_time.IndexOf(" "), vivod_energomera[k].t_time.Length - vivod_energomera[k].t_time.IndexOf(" ")).Trim();
                        asc[j].v(timeparser(streng), vivod_energomera[k].v_value);

                    }
                }
            }

            for (int i = 0; i < vivod_energomera_v.Count - 1; i++)
            {
                if (ls2.IndexOf(vivod_energomera_v[i].name_f) < 0)
                {
                    ls2.Add(vivod_energomera_v[i].name_f);
                    class4_1 m = new class4_1();
                    m.name_rp = vivod_energomera_v[i].name_f;
                    asc_v.Add(m);
                }
            }

            for (int k = 0; k < vivod_energomera_v.Count; k++)
            {

                for (int j = 0; j < asc_v.Count; j++)
                {

                    if (vivod_energomera_v[k].name_f.Equals(asc_v[j].name_rp))
                    {

                        string streng1 = vivod_energomera_v[k].t_time.Substring(vivod_energomera_v[k].t_time.IndexOf(" "), vivod_energomera_v[k].t_time.Length - vivod_energomera_v[k].t_time.IndexOf(" ")).Trim();
                        asc_v[j].v(timeparser(streng1), vivod_energomera_v[k].v_value);
                    }
                }
            }
            dgv.DataSource = null;
            dgv.DataSource = asc;
            dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
            dgv.Columns[0].HeaderText = "Название";  //присваиваем текст к заголовку колонок
            dgv.Columns[1].HeaderText = "00:00";
            dgv.Columns[2].HeaderText = "00:30";
            dgv.Columns[3].HeaderText = "01:00";
            dgv.Columns[4].HeaderText = "01:30";
            dgv.Columns[5].HeaderText = "02:00";
            dgv.Columns[6].HeaderText = "02:30";
            dgv.Columns[7].HeaderText = "03:00";
            dgv.Columns[8].HeaderText = "03:30";
            dgv.Columns[9].HeaderText = "04:00";
            dgv.Columns[10].HeaderText = "04:30";
            dgv.Columns[11].HeaderText = "05:00";
            dgv.Columns[12].HeaderText = "05:30";
            dgv.Columns[13].HeaderText = "06:00";
            dgv.Columns[14].HeaderText = "06:30";
            dgv.Columns[15].HeaderText = "07:00";
            dgv.Columns[16].HeaderText = "07:30";
            dgv.Columns[17].HeaderText = "08:00";
            dgv.Columns[18].HeaderText = "08:30";
            dgv.Columns[19].HeaderText = "09:00";
            dgv.Columns[20].HeaderText = "09:30";
            dgv.Columns[21].HeaderText = "10:00";
            dgv.Columns[22].HeaderText = "10:30";
            dgv.Columns[23].HeaderText = "11:00";
            dgv.Columns[24].HeaderText = "11:30";
            dgv.Columns[25].HeaderText = "12:00";
            dgv.Columns[26].HeaderText = "12:30";
            dgv.Columns[27].HeaderText = "13:00";
            dgv.Columns[28].HeaderText = "13:30";
            dgv.Columns[29].HeaderText = "14:00";
            dgv.Columns[30].HeaderText = "14:30";
            dgv.Columns[31].HeaderText = "15:00";
            dgv.Columns[32].HeaderText = "15:30";
            dgv.Columns[33].HeaderText = "16:00";
            dgv.Columns[34].HeaderText = "16:30";
            dgv.Columns[35].HeaderText = "17:00";
            dgv.Columns[36].HeaderText = "17:30";
            dgv.Columns[37].HeaderText = "18:00";
            dgv.Columns[38].HeaderText = "18:30";
            dgv.Columns[39].HeaderText = "19:00";
            dgv.Columns[40].HeaderText = "19:30";
            dgv.Columns[41].HeaderText = "20:00";
            dgv.Columns[42].HeaderText = "20:30";
            dgv.Columns[43].HeaderText = "21:00";
            dgv.Columns[44].HeaderText = "21:30";
            dgv.Columns[45].HeaderText = "22:00";
            dgv.Columns[46].HeaderText = "22:30";
            dgv.Columns[47].HeaderText = "23:00";
            dgv.Columns[48].HeaderText = "23:30";
            dgv.Columns[49].HeaderText = "24:00";


            dgv_vrem.DataSource = null;
            dgv_vrem.DataSource = asc_v;
            dgv_vrem.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
            dgv_vrem.Columns[0].HeaderText = "Название";  //присваиваем текст к заголовку колонок
            dgv_vrem.Columns[1].HeaderText = "00:00";
            dgv_vrem.Columns[2].HeaderText = "00:30";
            dgv_vrem.Columns[3].HeaderText = "01:00";
            dgv_vrem.Columns[4].HeaderText = "01:30";
            dgv_vrem.Columns[5].HeaderText = "02:00";
            dgv_vrem.Columns[6].HeaderText = "02:30";
            dgv_vrem.Columns[7].HeaderText = "03:00";
            dgv_vrem.Columns[8].HeaderText = "03:30";
            dgv_vrem.Columns[9].HeaderText = "04:00";
            dgv_vrem.Columns[10].HeaderText = "04:30";
            dgv_vrem.Columns[11].HeaderText = "05:00";
            dgv_vrem.Columns[12].HeaderText = "05:30";
            dgv_vrem.Columns[13].HeaderText = "06:00";
            dgv_vrem.Columns[14].HeaderText = "06:30";
            dgv_vrem.Columns[15].HeaderText = "07:00";
            dgv_vrem.Columns[16].HeaderText = "07:30";
            dgv_vrem.Columns[17].HeaderText = "08:00";
            dgv_vrem.Columns[18].HeaderText = "08:30";
            dgv_vrem.Columns[19].HeaderText = "09:00";
            dgv_vrem.Columns[20].HeaderText = "09:30";
            dgv_vrem.Columns[21].HeaderText = "10:00";
            dgv_vrem.Columns[22].HeaderText = "10:30";
            dgv_vrem.Columns[23].HeaderText = "11:00";
            dgv_vrem.Columns[24].HeaderText = "11:30";
            dgv_vrem.Columns[25].HeaderText = "12:00";
            dgv_vrem.Columns[26].HeaderText = "12:30";
            dgv_vrem.Columns[27].HeaderText = "13:00";
            dgv_vrem.Columns[28].HeaderText = "13:30";
            dgv_vrem.Columns[29].HeaderText = "14:00";
            dgv_vrem.Columns[30].HeaderText = "14:30";
            dgv_vrem.Columns[31].HeaderText = "15:00";
            dgv_vrem.Columns[32].HeaderText = "15:30";
            dgv_vrem.Columns[33].HeaderText = "16:00";
            dgv_vrem.Columns[34].HeaderText = "16:30";
            dgv_vrem.Columns[35].HeaderText = "17:00";
            dgv_vrem.Columns[36].HeaderText = "17:30";
            dgv_vrem.Columns[37].HeaderText = "18:00";
            dgv_vrem.Columns[38].HeaderText = "18:30";
            dgv_vrem.Columns[39].HeaderText = "19:00";
            dgv_vrem.Columns[40].HeaderText = "19:30";
            dgv_vrem.Columns[41].HeaderText = "20:00";
            dgv_vrem.Columns[42].HeaderText = "20:30";
            dgv_vrem.Columns[43].HeaderText = "21:00";
            dgv_vrem.Columns[44].HeaderText = "21:30";
            dgv_vrem.Columns[45].HeaderText = "22:00";
            dgv_vrem.Columns[46].HeaderText = "22:30";
            dgv_vrem.Columns[47].HeaderText = "23:00";
            dgv_vrem.Columns[48].HeaderText = "23:30";
            dgv_vrem.Columns[49].HeaderText = "24:00";

        } //небаланс по суткам для энергомеры
        public DataGridView report_neotvet(DataGridView dgv, ComboBox cb_rp, CheckedListBox clb, DateTimePicker dtp_n, DateTimePicker dtp_k, ComboBox cb_vrem_n, ComboBox cb_vrem_k, RichTextBox rtb, string punkt)
        {
            string z = "";
            dgv.DataSource = null;
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            result.Clear();

            //List<string> list_rp = new List<string>();

            //NpgsqlConnection connection = new NpgsqlConnection(monitor);
            //DataTable table = new DataTable();
            //connection.Open();

            //NpgsqlCommand com_rp = new NpgsqlCommand("select \"Name\" from \"GROUPS\" where \"Parent\" is null order by regexp_replace(\"Name\", '[^0-9]', '', 'g')::numeric ", connection);
            //NpgsqlDataReader datar = com_rp.ExecuteReader();
            //while (datar.Read()) { list_rp.Add(datar.GetString(0)); }

            ////MessageBox.Show(list_rp.Count().ToString());

            ////MessageBox.Show(list1.Count().ToString());

            //for (int t = 0; t < list_rp.Count; t++)
            //{

            //    // MessageBox.Show(list_rp[t].ToString());
            //    z = "select t4.\"Values_ID\", t0.\"Name\", t4.\"Name\", t4.\"Tag\" from (select t1.\"Groups_ID\", t2.\"Name\" from \"GROUPS\" t1, \"GROUPS\" t2 where t1.\"Tag\" = 'ТИИ' and t1.\"Parent\"=t2.\"Groups_ID\" and t2.\"Name\" = '" + list_rp[t].ToString() + "') as t0, \"TB_VALUES\" t4 where t4.\"Groups_ID\"=t0.\"Groups_ID\" and lower(t4.\"Name\") like 'пок%'";
            //    //z = "select t4.\"Values_ID\", t0.\"Name\", t4.\"Name\", t4.\"Tag\" from (select t1.\"Groups_ID\", t2.\"Name\" from \"GROUPS\" t1, \"GROUPS\" t2 where t1.\"Tag\" = 'ТИИ' and t1.\"Parent\"=t2.\"Groups_ID\" and t2.\"Name\" = '" + cb_rp.SelectedItem.ToString() + "') as t0, \"TB_VALUES\" t4 where t4.\"Groups_ID\"=t0.\"Groups_ID\" and t4.\"Name\" in  (" + tii(clb) + ")";
            //    //MessageBox.Show(z);
            //    NpgsqlCommand com = new NpgsqlCommand(z, connection);
            //    rtb.Text = rtb.Text + "\n \n" + z + "\n \n";
            //    DataTable table11 = new DataTable();
            //    NpgsqlDataAdapter dap = new NpgsqlDataAdapter(com);
            //    dap.Fill(table11);
            //    connection.Close();
            //    list1.Clear();
            //    foreach (DataRow row in table11.Rows)
            //    {
            //        class2 m = new class2();
            //        m.values_id = row.ItemArray[0].ToString();
            //        m.name = row.ItemArray[1].ToString();
            //        m.name_f = row.ItemArray[2].ToString();
            //        m.tag = row.ItemArray[3].ToString();
            //        list1.Add(m);
            //    }

            //MessageBox.Show(list1.Count().ToString());



            NpgsqlConnection connection = new NpgsqlConnection(monitor_options);
            connection.Open();
            if (punkt == "Все") { z = "select \"id\", \"name_rp\", \"name_tii\" from neotvet"; }
            else
            {
                z = "select \"id\", \"name_rp\", \"name_tii\" from neotvet where \"" + punkt + "\" = true";
            }

            NpgsqlCommand com = new NpgsqlCommand(z, connection);
            DataTable table = new DataTable();
            NpgsqlDataAdapter dap = new NpgsqlDataAdapter(com);
            dap.Fill(table);
            connection.Close();
            list1.Clear();
            foreach (DataRow row in table.Rows)
            {
                class2 m = new class2();
                m.values_id = row.ItemArray[0].ToString();
                m.name = row.ItemArray[1].ToString();
                m.name_f = row.ItemArray[2].ToString();
                //m.tag = row.ItemArray[3].ToString();
                list1.Add(m);
            }

            StringBuilder v_id = new StringBuilder();

            v_id.Append("(");
            for (int i = 0; i < list1.Count; i++)
            {

                v_id.Append(list1[i].values_id);
                if (i != list1.Count - 1)
                {
                    v_id.Append(", ");
                }
            }
            v_id.Append(")");

            DateTime d1 = new DateTime(dtp_n.Value.Year, dtp_n.Value.Month, dtp_n.Value.Day); //дата начала периода
            double x1 = d1.ToOADate();    //перевод нормальной даты в абсолютную

            DateTime d2 = new DateTime(dtp_k.Value.Year, dtp_k.Value.Month, dtp_k.Value.Day); //дата конца периода
            double x2 = d2.ToOADate();    //перевод нормальной даты в абсолютную

            SqlConnection con = new SqlConnection(sql_connection);
            con.Open();

            string chas1 = cb_vrem_n.SelectedItem.ToString();
            chas1 = chas1.Substring(0, chas1.IndexOf(":"));
            string min1 = cb_vrem_n.SelectedItem.ToString();
            min1 = min1.Substring(min1.IndexOf(":") + 1, 2);
            double ab_vr1 = (Convert.ToDouble(chas1) * 60 + Convert.ToDouble(min1)) * 0.000694443; //перевод часов и минут (то есть время) в абсолютное. Это от

            string chas2 = cb_vrem_k.SelectedItem.ToString();
            chas2 = chas2.Substring(0, chas2.IndexOf(":"));
            string min2 = cb_vrem_k.SelectedItem.ToString();
            min2 = min2.Substring(min2.IndexOf(":") + 1, 2);
            double ab_vr2 = (Convert.ToDouble(chas2) * 60 + Convert.ToDouble(min2)) * 0.000694445;  //перевод часов и минут (то есть время) в абсолютное. Это до

            double vr1 = x1 + ab_vr1;
            double vr2 = x2 + ab_vr2;
            string vr11 = Convert.ToString(vr1).Replace(",", ".");
            string vr22 = Convert.ToString(vr2).Replace(",", ".");

            if (vr1 > vr2) { MessageBox.Show("Неверный интервал времени."); }

            string z1 = "select VALUES_ID, T_TIME,  V_VALUE, VALID from ti where T_TIME >= " + vr11 + " and T_TIME <= " + vr22 + " and values_id in " + v_id.ToString() + " and VALID = '0' order by VALUES_ID, T_TIME ";
            SqlDataAdapter da = new SqlDataAdapter(z1, con);

            //MessageBox.Show("select VALUES_ID, T_TIME,  V_VALUE, VALID from ti where T_TIME >= " + vr11 + " and T_TIME <= " + vr22 + " and values_id in " + v_id.ToString() + " order by VALUES_ID, T_TIME ");

            rtb.Text = rtb.Text + z1 + "\n \n";
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataTable table1 = new DataTable();
            try
            {
                da.Fill(table1);
            }
            catch { }
            list.Clear();
            foreach (DataRow row in table1.Rows)
            {
                class0 m = new class0();
                m.t_time = row.ItemArray[1].ToString();
                m.values_id = row.ItemArray[0].ToString();
                m.v_value = row.ItemArray[2].ToString();
                m.valid = row.ItemArray[3].ToString();
                list.Add(m);
            }
            con.Close();

            for (int i = 0; i < list.Count; i++)
            {
                double x = Convert.ToDouble(list[i].t_time);
                DateTime dt = DateTime.FromOADate(x);
                list[i].t_time = dt.ToString();
            }

            //NpgsqlConnection connection1 = new NpgsqlConnection(monitor_options);
            //connection1.Open();
            //NpgsqlCommand com1 = new NpgsqlCommand("select \"name_tii\", \"ktt\"  from \"balans\" where \"name_rp\" = '" + rp + "' ", connection1);
            //rtb.Text = rtb.Text + "select \"name_tii\", \"ktt\"  from \"balans\" where \"name_rp\" = '" + rp + "' \n \n";
            //NpgsqlDataReader datar1 = com1.ExecuteReader();
            //fider.Clear();
            //while (datar1.Read())
            //{
            //    class5 m = new class5();
            //    m.fider = datar1.GetValue(0).ToString();
            //    m.val = datar1.GetValue(1).ToString();
            //    fider.Add(m);

            //}
            //connection1.Close();
            //schet.Clear();
            //NpgsqlConnection connection2 = new NpgsqlConnection(monitor_options);
            //connection2.Open();
            //NpgsqlCommand com2 = new NpgsqlCommand("select \"name_tii\", \"num_sch\"  from \"number_sch\" where \"name_rp\" = '" + rp + "'", connection2);
            //rtb.Text = rtb.Text + "select \"name_tii\", \"num_sch\"  from \"number_sch\" where \"name_rp\" = '" + rp + "' \n \n";
            //NpgsqlDataReader datar2 = com2.ExecuteReader();
            ////connection2.Close();

            //while (datar2.Read())
            //{
            //    class6 m = new class6();
            //    m.name_tii = datar2.GetValue(0).ToString();
            //    m.num_schet = datar2.GetValue(1).ToString();
            //    schet.Add(m);

            //}
            //connection2.Close();

            for (int i = 0; i < list.Count; i++)
            {
                for (int j = 0; j < list1.Count; j++)
                {
                    if (list[i].values_id == list1[j].values_id)
                    {
                        class3 m = new class3();
                        m.name = list1[j].name;
                        m.name_f = list1[j].name_f;
                        m.tag = list1[j].tag;
                        m.t_time = list[i].t_time;
                        m.v_value = list[i].v_value;
                        m.valid = list[i].valid;
                        //m.meas = list1[j].meas;
                        result.Add(m);
                    }
                }
            }

            //}
            dgv.DataSource = null;  //обнуляем таблицу
            dgv.DataSource = result; //заполняем таблицу листом result
            //for (int i = 0; i < dgv.Rows.Count; i++)
            //{
            //    for (int j = 0; j < dgv.Columns.Count; j++)
            //    {
            //        for (int k = 0; k < fider.Count; k++)
            //        {
            //            if (dgv.Rows[i].Cells[1].Value.ToString().Equals(fider[k].fider)) { dgv.Rows[i].Cells[4].Value = fider[k].val; }
            //        }
            //        for (int n = 0; n < schet.Count; n++)
            //        {

            //            if (dgv.Rows[i].Cells[1].Value.ToString().Equals(schet[n].name_tii)) { dgv.Rows[i].Cells[3].Value = schet[n].num_schet; }
            //            //if (dgv.Rows[i].Cells[1].Value.ToString().IndexOf(' ', dgv.Rows[i].Cells[1].Value.ToString().Length).Equals(schet[n].name_tii.IndexOf(' ', schet[n].name_tii.Length))) { dgv.Rows[i].Cells[3].Value = schet[n].num_schet; }

            //        }
            //    }
            //}
            dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgv.Columns[2].Visible = false;  //скрываем столбец с названием ТИИ
            dgv.Columns[0].HeaderText = "Объект";
            dgv.Columns[1].HeaderText = "Присоединение"; //присваиваем текст к заголовку колонок
            dgv.Columns[3].Visible = false;
            dgv.Columns[4].Visible = false;
            dgv.Columns[5].HeaderText = "Дата и время";
            dgv.Columns[6].HeaderText = "Значение";
            dgv.Columns[7].HeaderText = "*";

            return dgv;

        }

        //Энергосфера

        public ComboBox spisok_r(ComboBox cb, RichTextBox rtb)
        {
            SqlConnection connection = new SqlConnection(put_sphera);
            cb.Items.Clear();
            string z = "Select PointName from Points where ID_Parent = 1 order by PointName";
            rtb.Text = rtb.Text + z + "\n";
            connection.Open();
            SqlCommand com = new SqlCommand(z, connection);
            SqlDataReader datar = com.ExecuteReader();
            while (datar.Read())
            {
                cb.Items.Add(datar.GetString(0));
            }
            connection.Close();
            cb.SelectedItem = "ПС Р-1";

            return cb;
        }
        public CheckedListBox spisok_othod_r(CheckedListBox clb, ComboBox cb, RichTextBox rtb)
        {
            SqlConnection connection = new SqlConnection(put_sphera);
            clb.Items.Clear();

            string z = "SELECT t0.PointName FROM Points t0 RIGHT JOIN " +
                "(SELECT ID_Point, PointName, ID_Parent, Point_Type, ID_Ref FROM Points WHERE PointName = '" + cb.SelectedItem.ToString() + "') as t1 ON t0.ID_Parent = t1.ID_Point";
            rtb.Text = rtb.Text + z + "\n";
            connection.Open();
            SqlCommand command = new SqlCommand(z, connection);
            SqlDataReader datar = command.ExecuteReader();
            while (datar.Read()) { clb.Items.Add(datar.GetString(0)); }
            connection.Close();

            return clb;
        }
        public DataGridView resultat_sphera(DataGridView dgv, ComboBox cb_r, CheckedListBox clb, DateTimePicker dtp_n, DateTimePicker dtp_k, ComboBox cb_vrem_n, ComboBox cb_vrem_k, RichTextBox rtb)
        {
            
            
            vivod_sphera.Clear();
            dgv.DataSource = null;

            DateTime d1 = new DateTime(dtp_n.Value.Year, dtp_n.Value.Month, dtp_n.Value.Day); //дата начала периода
            //double x1 = d1.ToOADate();    //перевод нормальной даты в абсолютную

            DateTime d2 = new DateTime(dtp_k.Value.Year, dtp_k.Value.Month, dtp_k.Value.Day); //дата конца периода
            //double x2 = d2.ToOADate();    //перевод нормальной даты в абсолютную

            //string chas1 = cb_vrem_n.SelectedItem.ToString();
            //chas1 = chas1.Substring(0, chas1.IndexOf(":"));
            //string min1 = cb_vrem_n.SelectedItem.ToString();
            //min1 = min1.Substring(min1.IndexOf(":") + 1, 2);
            //double ab_vr1 = (Convert.ToDouble(chas1) * 60 + Convert.ToDouble(min1)) * 0.000694443; //перевод часов и минут (то есть время) в абсолютное. Это от

            //string chas2 = cb_vrem_k.SelectedItem.ToString();
            //chas2 = chas2.Substring(0, chas2.IndexOf(":"));
            //string min2 = cb_vrem_k.SelectedItem.ToString();
            //min2 = min2.Substring(min2.IndexOf(":") + 1, 2);
            //double ab_vr2 = (Convert.ToDouble(chas2) * 60 + Convert.ToDouble(min2)) * 0.000694445;  //перевод часов и минут (то есть время) в абсолютное. Это до

            //double vr1 = x1 + ab_vr1;
            //double vr2 = x2 + ab_vr2;

            //string n_d = DateTime.FromOADate(vr1).ToString().Replace(".", "-");
            //string k_d = DateTime.FromOADate(vr2).ToString().Replace(".", "-");

            SqlConnection connection = new SqlConnection(put_sphera);
            DataTable table = new DataTable();
            
            //string z = "SELECT t8.PointName, Max(t9.Coeff), t8.DT, MAX(t8.Val), Max(t9.Coeff)*MAX(t8.Val)   from ASCUE.dbo.SchemaContents t9 " +
            //"RIGHT JOIN (SELECT t6.ID_PP, t6.DT, t6.Val, t7.PointName  FROM PointMains t6 " +
            //"RIGHT JOIN (SELECT t4.ID_PP, t5.PointName FROM ASCUE.dbo.PointParams t4 " +
            //"RIGHT JOIN (SELECT t2.ID_Point, t3.PointName, t2.ID_Parent FROM ASCUE.dbo.Points t2 " +
            //"RIGHT JOIN (SELECT t0.ID_Point, t0.PointName, t0.ID_Parent FROM ASCUE.dbo.Points t0 " +
            //"RIGHT JOIN (SELECT ID_Point, PointName, ID_Parent, Point_Type, ID_Ref FROM ASCUE.dbo.Points WHERE PointName = '" + cb_r.SelectedItem.ToString() + "') as t1 " +
            //"ON t0.ID_Parent = t1.ID_Point) as t3 " +
            //"ON t2.ID_Parent = t3.ID_Point where t2.Point_Type = 21) as t5 " +
            //"ON t4.ID_Point = t5.ID_Point where t4.ID_Param = 4 AND t5.PointName in (" + tii(clb) + ")) as t7 " +
            //"ON t6.ID_PP = t7.ID_PP and DT>= '" + n_d + "' and DT<='" + k_d + "') as t8 ON t9.ID_Ref=t8.ID_PP GROUP BY t8.PointName, t8.DT ORDER BY t8.PointName"; //Запрос для мощности

            string z = "SELECT t8.PointName, Max(t9.Coeff), t8.DT, MAX(t8.Val) from ASKUE.dbo.SchemaContents t9 " +
            "RIGHT JOIN (SELECT t7.ID_PP, t7.PointName, t6.DT, t6.Val FROM ASKUE.dbo.PointNIs_On_Main_Stack t6 " + 
            "RIGHT JOIN (SELECT t4.ID_PP, t5.PointName FROM ASKUE.dbo.PointParams t4 " +
            "RIGHT JOIN (SELECT t2.ID_Point, t3.PointName, t2.ID_Parent FROM ASKUE.dbo.Points t2 " +
            "RIGHT JOIN (SELECT t0.ID_Point, t0.PointName, t0.ID_Parent FROM ASKUE.dbo.Points t0 " +
            "RIGHT JOIN (SELECT ID_Point, PointName, ID_Parent, Point_Type, ID_Ref FROM ASKUE.dbo.Points WHERE PointName = '" + cb_r.SelectedItem.ToString() + "') as t1 " +
            "ON t0.ID_Parent = t1.ID_Point where t0.PointName in ("+ tii(clb) +"))  as t3 " +
            "ON t2.ID_Parent = t3.ID_Point where t2.Point_Type = 21) as t5 " +
            "ON t4.ID_Point = t5.ID_Point where t4.ID_Param = 4) as t7 " +
            "ON t6.ID_PP=t7.ID_PP where t6.DT>= '" + d1 + "' and t6.DT<='" + d2 + "') as t8 " + 
            "ON t9.ID_Ref=t8.ID_PP GROUP BY t8.PointName, t8.DT ORDER BY t8.PointName";   //Запрос для показаний
            
            rtb.Text = rtb.Text + z + "\n\n";
            connection.Open();
            SqlCommand com = new SqlCommand(z, connection);
            SqlDataAdapter dap = new SqlDataAdapter(com);
            dap.Fill(table);
            connection.Close();

            foreach (DataRow row in table.Rows)
            {
                res_sphera m = new res_sphera();
                m.name_r = cb_r.SelectedItem.ToString().Substring(cb_r.SelectedItem.ToString().IndexOf("ПС ") + 2);
                m.name_l = row.ItemArray[0].ToString().Substring(row.ItemArray[0].ToString().IndexOf("Л"));
                m.koef = row.ItemArray[1].ToString();
                m.time = row.ItemArray[2].ToString();
                m.val_pok = row.ItemArray[3].ToString();
                vivod_sphera.Add(m);
            }
            if (vivod_sphera.Count == 0)
            {
                MessageBox.Show("Нет данных! Проверьте дату.");
            }
            else { dgv.DataSource = vivod_sphera; }
            
            dgv.Columns[0].HeaderText = "Объект";
            dgv.Columns[1].HeaderText = "Присоединение";
            dgv.Columns[2].HeaderText = "Кт";
            dgv.Columns[3].HeaderText = "Дата и время";
            dgv.Columns[4].HeaderText = "Показания";

            dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True;

            return dgv;
        }

        //Меркурии

        public ComboBox spisok_tp_mercury(ComboBox cb)
        {
            cb.Items.Clear();
            cb.Items.Add("ВРЭС");
            cb.Items.Add("ЗРЭС");
            cb.Items.Add("СРЭС");
            cb.Items.Add("ЮРЭС");
            cb.SelectedItem = "ВРЭС";

            return cb;
        }
        public CheckedListBox spisok_othod_tp_mercury(CheckedListBox clb, ComboBox cb, ComboBox cb_base, RichTextBox rtb)
        {
            string res = "";
            //if (cb.SelectedItem.ToString() == "ВРЭС") { res = "tree_device_mercury_v"; }
            //if (cb.SelectedItem.ToString() == "ЗРЭС") { res = "tree_device_mercury_z"; }
            //if (cb.SelectedItem.ToString() == "СРЭС") { res = "tree_device_mercury_s"; }
            //if (cb.SelectedItem.ToString() == "ЮРЭС") { res = "tree_device_mercury_u"; }
            if (cb.SelectedItem.ToString() == "ВРЭС") { res = "\"RES_treemenuv\""; }
            if (cb.SelectedItem.ToString() == "ЗРЭС") { res = "\"RES_treemenuz\""; }
            if (cb.SelectedItem.ToString() == "СРЭС") { res = "\"RES_treemenus\""; }
            if (cb.SelectedItem.ToString() == "ЮРЭС") { res = "\"RES_treemenuu\""; }
            string z = "select upper(name) from " + res + " where name like 'ТП%' order by name";
            rtb.Text = rtb.Text + z + "\n";
       
            NpgsqlConnection connection = new NpgsqlConnection(mercury);
            clb.Items.Clear();
            connection.Open();
            NpgsqlCommand command = new NpgsqlCommand(z, connection);
            NpgsqlDataReader datar = command.ExecuteReader();
            while (datar.Read())
            {
                clb.Items.Add(datar.GetString(0));
            }
            connection.Close();
            return clb;
        }
        public DataGridView resultat_mercury(DataGridView dgv, ComboBox cb_res, CheckedListBox clb, DateTimePicker dtp_n, DateTimePicker dtp_k, RichTextBox rtb)
        {
            string res = "";
            DateTime d1 = new DateTime(dtp_n.Value.Year, dtp_n.Value.Month, dtp_n.Value.Day); 
            DateTime d2 = new DateTime(dtp_k.Value.Year, dtp_k.Value.Month, dtp_k.Value.Day);
            if (cb_res.SelectedItem.ToString() == "ВРЭС") { res = "v"; }
            if (cb_res.SelectedItem.ToString() == "ЗРЭС") { res = "z"; }
            if (cb_res.SelectedItem.ToString() == "СРЭС") { res = "s"; }
            if (cb_res.SelectedItem.ToString() == "ЮРЭС") { res = "u"; }

          //  string z = "select t7.name, t8.name, t8.serial_number, t8.date, t8.energy_reset_sum from tree_device_mercury_"+res +" t7 " +
          //"right join (select t6.parent_id, t5.energy_day, t5.energy_reset_sum, t5.power_day, t5.date, t6.name, t6.serial_number from data_mercury_"+res+" t5 " +
          //"right join (select t3.serial_number, t4.name, t4.id, t4.parent_id from device_mercury_"+res+" t3 " +
          //"right join (select t2.id, t2.parent_id, t2.name from tree_device_mercury_"+res+" t1, tree_device_mercury_"+res+" t2 where t1.name in("+tii(clb)+") and t1.id=t2.parent_id) as t4 " + 
          //"on t3.id = t4.id) as t6 " +
          //"on t5.serial_number = t6.serial_number where t5.date between '" + dtp_n.Value.ToString("yyyy-MM-dd") + "' and '" + dtp_k.Value.ToString("yyyy-MM-dd") + "') as t8 " +
          //"on t7.id = t8.parent_id order by t7.name, t8.name, t8.date";

            string z = "select t7.name, t8.name, t8.serial_number, t8.date, t8.energy_reset_sum from \"RES_treemenu" + res + "\" t7 " +
          "right join (select t6.parent_id, t5.energy_day, t5.energy_reset_sum, t5.power_day, t5.date, t6.name, t6.serial_number from data_mercury_" + res + " t5 " +
          "right join (select t3.serial_number, t4.name, t4.id, t4.parent_id from device_mercury_" + res + " t3 " +
          "right join (select t2.id, t2.parent_id, t2.name from \"RES_treemenu" + res + "\" t1, \"RES_treemenu" + res + "\" t2 where t1.name in(" + tii(clb) + ") and t1.id=t2.parent_id) as t4 " +
          "on t3.id = t4.id) as t6 " +
          "on t5.serial_number = t6.serial_number where t5.date between '" + dtp_n.Value.ToString("yyyy-MM-dd") + "' and '" + dtp_k.Value.ToString("yyyy-MM-dd") + "') as t8 " +
          "on t7.id = t8.parent_id order by t7.name, t8.name, t8.date";
            rtb.Text = rtb.Text + z + "\n";

            NpgsqlConnection connection = new NpgsqlConnection(mercury);
            DataTable table = new DataTable();
            connection.Open();
            NpgsqlCommand com = new NpgsqlCommand(z, connection);
            NpgsqlDataAdapter dap = new NpgsqlDataAdapter(com);
            dap.Fill(table);
            connection.Close();
            vivod_mercury.Clear();
            foreach (DataRow row in table.Rows)
            {
                res_mercury m = new res_mercury();
                m.name_tp = row.ItemArray[0].ToString();
                m.name_t = row.ItemArray[1].ToString();
                m.s_number = row.ItemArray[2].ToString();
                m.data = row.ItemArray[3].ToString();
                m.tek_val = row.ItemArray[4].ToString();
                vivod_mercury.Add(m);
            }

            dgv.DataSource = null;
            dgv.DataSource = vivod_mercury;
            dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells); //автоматический размер колонок
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgv.Columns[0].HeaderText = "ТП";
            dgv.Columns[1].HeaderText = "Т";
            dgv.Columns[2].HeaderText = "Номер сч";
            dgv.Columns[3].HeaderText = "Дата";
            dgv.Columns[4].HeaderText = "Показания";

            return dgv;
        }

        //Энергомера ТП

        public ComboBox spisok_tp_energo(ComboBox cb, CheckBox ch_b_v, CheckBox ch_b_z, CheckBox ch_b_s, CheckBox ch_b_u, RichTextBox rtb)
        {
            string res = "В";
            if (ch_b_v.Checked == true) { res = "В"; }
            if (ch_b_z.Checked == true) { res = "З"; }
            if (ch_b_s.Checked == true) { res = "С"; }
            if (ch_b_u.Checked == true) { res = "Ю"; }
            string z = "select \"Name\" from \"ComObjects\" where \"Name\" like '" + res + " ТП%' order by \"Name\" ";
            rtb.Text = rtb.Text + z + "\n";
            NpgsqlConnection connection = new NpgsqlConnection(energomera_tp_connection);
            cb.Items.Clear();
            connection.Open();
            NpgsqlCommand com = new NpgsqlCommand(z, connection);
            NpgsqlDataReader datar = com.ExecuteReader();
            while (datar.Read()) { cb.Items.Add(datar.GetString(0)); }
            cb.SelectedItem = cb.Items[0].ToString();
            connection.Close();
            return cb;
        } //список тп для энергомеры
        public CheckedListBox spisok_othod_tp_energo(CheckedListBox clb, ComboBox cb, RichTextBox rtb)
        {
            NpgsqlConnection connection = new NpgsqlConnection(energomera_tp_connection);
            clb.Items.Clear();
            connection.Open();
            string z = "select \"SerialNumber\" from \"Meters\" where \"ComObjectId\" IN (select \"ComObjectId\" from \"ComObjects\" where \"Name\" = '"+cb.SelectedItem.ToString()+"')";
            rtb.Text = rtb.Text + z + "\n";
            NpgsqlCommand command = new NpgsqlCommand(z, connection);
            NpgsqlDataReader datar = command.ExecuteReader();
            while (datar.Read()) { clb.Items.Add(datar.GetString(0)); }
            connection.Close();
            return clb;

        } //список отходящих для энергомеры ТП
        public DataGridView resultat_tp_energo(DataGridView dgv, ComboBox cb_tp, CheckedListBox clb, DateTimePicker dtp_n, DateTimePicker dtp_k, RichTextBox rtb)
        {

            DateTime d1 = new DateTime(dtp_n.Value.Year, dtp_n.Value.Month, dtp_n.Value.Day);
            DateTime d2 = new DateTime(dtp_k.Value.Year, dtp_k.Value.Month, dtp_k.Value.Day);
            DateTime d3 = new DateTime(dtp_n.Value.Year, dtp_n.Value.Month, dtp_n.Value.Day);
            DateTime d4 = new DateTime(dtp_k.Value.Year, dtp_k.Value.Month, dtp_k.Value.Day);
            d3 = d3.AddDays(1);
            d4 = d4.AddDays(1);
            string z = "select t2.\"SerialNumber\", t3.\"DT\", max(t3.\"Val\") from \"Values\" t3 " +
            "right join (SELECT t1.\"MeterId\", t1.\"SerialNumber\" FROM \"Meters\" t1 where t1.\"SerialNumber\" in (" + tii(clb) + ")) as t2 " +
            "on t3.\"MeterId\" = t2.\"MeterId\" where \"TariffId\" = 1 and t3.\"DT\" BETWEEN '" + d1 + "' and '" + d2 + "' group by t2.\"MeterId\", t2.\"SerialNumber\", t3.\"DT\"";
            rtb.Text = rtb.Text + z + "\n";
//            string z = "select t5.\"SerialNumber\",t5.\"DT\", max(t5.\"Val\"), max(t4.\"Val\")  from \"Values\" t4 " +
// "right join (select t2.\"MeterId\", t2.\"SerialNumber\", t3.\"DT\", t3.\"Val\" from \"Values\" t3 " +
//"right join (SELECT t1.\"MeterId\", t1.\"SerialNumber\" FROM \"Meters\" t1 where t1.\"SerialNumber\" in ("+ tii(clb) +" )) as t2 " +
//"on t3.\"MeterId\" = t2.\"MeterId\" where \"TariffId\" = 1 and t3.\"DT\" BETWEEN '"+d1+"' and '"+d2+"' ) as t5 " +
//"on t4.\"MeterId\" = t5.\"MeterId\" where \"TariffId\" = 1 and t4.\"DT\" BETWEEN '"+d3 +"' and '"+d4+"' group by t5.\"MeterId\", t5.\"SerialNumber\", t5.\"DT\"";

            NpgsqlConnection connection = new NpgsqlConnection(energomera_tp_connection);
            DataTable table = new DataTable();
            connection.Open();
            NpgsqlCommand com = new NpgsqlCommand(z, connection);
            NpgsqlDataAdapter dap = new NpgsqlDataAdapter(com);
            dap.Fill(table);
            connection.Close();
            vivod_tp_energo.Clear();
             foreach (DataRow row in table.Rows)
                {
                 res_tp_energo m = new res_tp_energo();
                 m.name_tp = cb_tp.SelectedItem.ToString();
                 m.s_number = row.ItemArray[0].ToString();
                 m.data = row.ItemArray[1].ToString();
                 m.val =  Math.Round(Convert.ToDecimal( row.ItemArray[2]),3).ToString();
                 vivod_tp_energo.Add(m); 
            }
           
            dgv.DataSource = null;
            dgv.DataSource = vivod_tp_energo;
            dgv.Columns[0].HeaderText = "ТП";
            dgv.Columns[1].HeaderText = "Номер сч.";
            dgv.Columns[2].HeaderText = "Дата";
            dgv.Columns[3].HeaderText = "Показания";
            dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells); //автоматический размер колонок
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                return dgv;
        }
        public void poisk_tp(TextBox tb, ComboBox cb, CheckedListBox clb)
        {
            if (tb.Text != "")
            {
                NpgsqlConnection conn = new NpgsqlConnection(energomera_tp_connection );
                string z = "select \"SerialNumber\" from \"Meters\" where \"ComObjectId\" IN (select \"ComObjectId\" from \"ComObjects\" where \"Name\" = '" + cb.SelectedItem.ToString() + "') and \"SerialNumber\" like '%" + tb.Text.ToString() + "%' ";
                
                conn.Open();
                NpgsqlCommand comm = new NpgsqlCommand(z, conn);
                NpgsqlDataReader dr = comm.ExecuteReader();
                clb.Items.Clear();
                while (dr.Read())
                {
                    for (int i = 0; i < dr.FieldCount; i++)
                    {
                        clb.Items.Add(dr[i].ToString());
                    }
                }

                conn.Close();

            }
        }

    }
}
