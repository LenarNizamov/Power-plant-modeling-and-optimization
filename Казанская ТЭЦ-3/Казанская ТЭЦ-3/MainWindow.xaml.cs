using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO; //Подключаем пространство имен для работы с файлами
using System.Globalization;
using System.Diagnostics;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace Казанская_ТЭЦ_3
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            StreamReader stream = new StreamReader("исходные данные.txt"); STR55 = File.ReadAllLines("исходные данные.txt");
            textBox1.Text = STR55[1]; textBox2.Text = STR55[2]; textBox3.Text = STR55[3]; textBox4.Text = STR55[4]; textBox14.Text = STR55[5]; textBox7.Text = STR55[6]; textBox28.Text = STR55[7]; textBox29.Text = STR55[8]; textBox30.Text = STR55[9];
            textBox70.Text = STR55[10]; textBox5.Text = STR55[11]; textBox8.Text = STR55[12]; textBox10.Text = STR55[13]; textBox9.Text = STR55[14]; textBox11.Text = STR55[15]; textBox12.Text = STR55[16]; textBox16.Text = STR55[17]; textBox39.Text = STR55[18]; textBox57.Text = STR55[19];
            textBox58.Text = STR55[20]; textBox59.Text = STR55[21]; textBox60.Text = STR55[22]; textBox61.Text = STR55[23]; textBox62.Text = STR55[24]; textBox63.Text = STR55[25]; textBox64.Text = STR55[26];          
        }

        public double N1, Qt1, D1t, Dp2, N3, Qt3, N4, Qt4, Dp5, N6, Qt6, Dp6, Qt31, Qt41, Qt61, D140_13, D140_30, Dp30_13, D3013, Qpb, Qpvk, Dptn1, Dptn2, lt, lp, ln, lp30, N5, N2, Dv6, Dv2, Dv51, Dv52, D30_13, Dv30_13, Dv140_13, D140_45, Qpsn, Ns, Qs, Qt, Gcw4, Gcw3, Gcw6, Gcw1, k13, k30, k45, kpb, P6, P4, P3, P1, Qp13,
        N1max, N1min, N3max, N3min, N4max, N4min, N6max, N6min, Qt6max, Qt4max, Qt3max, Qt1max, Qt6min, Qt4min, Qt3min, Qt1min, Qbmax, Dp5max, Dp5min, Dp6max, D30_13max, Dp2min, Dp2max, Q30_13min, D140_13max, Dp30, tcw1, tcw4, tcw3, tcw6, tcwo3, tcwo4, tcwo6, Pchrp, k1, k2, k3, k4, k5, k6, k7, Gek, Ghob, Gsw, Gsetw, Gmh, G13pr, Gdist, Gpts, tsw,
        Qpm13, G1p, G2p, G3p, G4p, G5p, G6p, G7p, S, nk, Nko, Ncn, Nto, anas, apvk, Qe, Qsnto, Ke, Nsnee, Nsn, B, Be, ntp, Kotrk, nnet, be, bte, dQe, Kotrt, Qsnko, Ct, Cn, Cnsn, bteh, Cgv, C13, C30, C45, Qpvkmax, Dptnmax, Dptnmin, t1, h1, hdop, hv, h13, h30, h45, h30n, hpp, Ddvd, Ddnd, Dpsv, lN31, lN32, lN11, lN12, lN13, lN41, lN42, lN43, lN44, lN45, lN61, lN62, lN63, Q1t,
        MMM, MMM1, MMM2, MM, a, dptn1, dptn2, tp1, tp3, tp4, tp6, np1, np3, np4, np6, dp2, dp5, dp6, dp3013, Fs, D, D11, rty, q, Gsetv1, Gsetv3, Gsetv4, Gsetv6, tos1, tos3, tos4, tos6, tnv, dtk1, dtk3, dtk4, dtk6, dtsp1, dtsp3, dtsp4, dtsp6, Nv1, Nv3, Nv4, Nv6, rtn, tvv, Qpodp, tpodp, Qknom;

        public int ogrN1, ogrN3, ogrN4, ogrN6, ogrQt1, ogrQt3, ogrQt4, ogrQt6, indtg1, indtg2, indtg3, indtg4, indtg5, indtg6, indk1, indk2, indk3, indk4, indk5, indk6, indk7, cmb4, cmb5,
        ch1, ch2, ch3, ch4, ch5, indcn1, indcn3, indcn4, indptn1, indptn2, indP3, indP4, indP6, in6, indG1, indG3, indG4, indG6;
        public int im = 0;
        public string iii1, iii2, iii3, iii4, iii5, iii6, kkk1, kkk2, kkk3, kkk4, kkk5, kkk6, kkk7, iii11, iii21, iii31, iii41, iii51, iii61, iiiptn1, iiiptn2, iiiptn11, iiiptn21, kkk11, kkk21, kkk31, kkk41, kkk51, kkk61, kkk71;

        const double G1min = 180; const double G2min = 180; const double G3min = 180; const double G4min = 180; const double G5min = 180; const double G6min = 180; const double G7min = 180;
        const double G1max = 400; const double G2max = 380; const double G3max = 370; const double G4max = 400; const double G5max = 480; const double G6max = 400; const double G7max = 400;

        double kp, kt; // коэф.ц.пр. и т. 

        public double[] N1M; public double[,] Ge;
        public int Nvar;
        Stopwatch sWatch = new Stopwatch();

        public Excel.Application excelapp = new Excel.Application();
        public Excel.Sheets excelsheets;
        public Excel.Worksheet excelworksheet;
        public Excel.Range excelcells;
        public Excel.Workbooks excelappworkbooks;
        public Excel.Workbook excelappworkbook;

        public double[,] ISHD = new double[31, 31];
        public double[] X = new double[35];
        public double[] Y = new double[100];
        public double[] F = new double[24];
        public double[] X2 = new double[100];
        public double[,] X3 = new double[128, 100];
        public double[] F1 = new double[24];
        public double[] Pr = new double[24];
        public double[] VAR = new double[100];
        public double[][] Ja = MatrixDecompositionProgram.MatrixCreate(24, 24);
        public string[] STR1; public string[] STR2; public string[] STR3; public string[] STR4; public string[] STR5; public string[] STR6; public string[] STR7; public string[] STR8; public string[] STR9; public string[] STR10;
        public string[] STR11; public string[] STR12; public string[] STR13; public string[] STR14; public string[] STR15; public string[] STR16; public string[] STR17; public string[] STR18; public string[] STR19; public string[] STR20;
        public string[] STR21; public string[] STR22; public string[] STR23; public string[] STR24; public string[] STR25; public string[] STR26; public string[] STR27; public string[] STR28; public string[] STR29; public string[] STR30;
        public string[] STR31; public string[] STR32; public string[] STR33; public string[] STR34; public string[] STR35; public string[] STR36; public string[] STR37; public string[] STR38; public string[] STR39; public string[] STR40;
        public string[] STR41; public string[] STR42; public string[] STR43; public string[] STR44; public string[] STR45; public string[] STR46; public string[] STR47; public string[] STR48; public string[] STR49; public string[] STR50;
        public string[] STR51; public string[] STR52; public string[] STR53; public string[] STR54; public string[] STR55; public string[] STR56; public string[] STR57; public string[] STR58; public string[] STR59; public string[] STR60;
        public string[] STR61; public string[] STR62; public string[] STR63; public string[] STR64; public string[] STR65; public string[] STR66; public string[] STR67; public string[] STR68; public string[] STR69; public string[] STR70;
        public string[] STR71;
        public double[] DSTR1; public double[] DSTR2; public double[] DSTR3; public double[] DSTR4; public double[] DSTR5; public double[] DSTR6; public double[] DSTR7; public double[] DSTR8; public double[] DSTR9; public double[] DSTR10;
        public double[] DSTR11; public double[] DSTR12; public double[] DSTR13; public double[] DSTR14; public double[] DSTR15; public double[] DSTR16; public double[] DSTR17; public double[] DSTR18; public double[] DSTR19; public double[] DSTR20;
        public double[] DSTR21; public double[] DSTR22; public double[] DSTR23; public double[] DSTR24; public double[] DSTR25; public double[] DSTR26; public double[] DSTR27; public double[] DSTR28; public double[] DSTR29; public double[] DSTR30;
        public double[] DSTR31; public double[] DSTR32; public double[] DSTR33; public double[] DSTR34; public double[] DSTR35; public double[] DSTR36; public double[] DSTR37; public double[] DSTR38; public double[] DSTR39; public double[] DSTR40;
        public double[] DSTR41; public double[] DSTR42; public double[] DSTR43; public double[] DSTR44; public double[] DSTR45; public double[] DSTR46; public double[] DSTR47; public double[] DSTR48; public double[] DSTR49; public double[] DSTR50;
        public double[] DSTR51; public double[] DSTR52; public double[] DSTR53; public double[] DSTR54; public double[] DSTR55; public double[] DSTR56; public double[] DSTR57; public double[] DSTR58; public double[] DSTR59; public double[] DSTR60;
        public double[] DSTR61; public double[] DSTR62; public double[] DSTR63; public double[] DSTR64; public double[] DSTR65; public double[] DSTR66; public double[] DSTR67; public double[] DSTR68; public double[] DSTR69; public double[] DSTR70;
        public double[] DSTR71;
                
        public void Button_Click_1(object sender, EventArgs e)
        {            
            sWatch.Start();

            STR55[0] = textBox6.Text; STR55[1] = textBox1.Text; STR55[2] = textBox2.Text; STR55[3] = textBox3.Text; STR55[4] = textBox4.Text; STR55[5] = textBox14.Text; STR55[6] = textBox7.Text; STR55[7] = textBox28.Text; STR55[8] = textBox29.Text; STR55[9] = textBox30.Text;
            STR55[10] = textBox70.Text; STR55[11] = textBox5.Text; STR55[12] = textBox8.Text; STR55[13] = textBox10.Text; STR55[14] = textBox9.Text; STR55[15] = textBox11.Text; STR55[16] = textBox12.Text; STR55[17] = textBox16.Text; STR55[18] = textBox39.Text; STR55[19] = textBox57.Text;
            STR55[20] = textBox58.Text; STR55[21] = textBox59.Text; STR55[22] = textBox60.Text; STR55[23] = textBox61.Text; STR55[24] = textBox62.Text; STR55[25] = textBox63.Text; STR55[26] = textBox64.Text;
            Replace();

            StreamWriter streamWriter = new StreamWriter("исходные данные.txt");
            for (int i = 0; i < 27; i++) { streamWriter.WriteLine(STR55[i]); }
            streamWriter.Close();
            combobox21.Items.Clear();
            if (combobox13.Text != "работа" & combobox14.Text != "работа" & combobox15.Text != "работа" & combobox16.Text != "работа" & combobox17.Text != "работа" & combobox18.Text != "работа" & combobox19.Text != "работа" & combobox20.Text == "запрещено" & combobox10.Text != "активен") combobox20.Text = "разрешено";
            if (combobox10.Text == "активен")
            {
                if (combobox1.Text == "работа") combobox1.Text = "резерв"; if (combobox2.Text == "работа") combobox2.Text = "резерв";
                if (combobox3.Text == "работа") combobox3.Text = "резерв"; if (combobox4.Text == "работа") combobox4.Text = "резерв";
                if (combobox5.Text == "работа") combobox5.Text = "резерв"; if (combobox6.Text == "работа") combobox6.Text = "резерв";
            }
            Stream();
            double.TryParse(textBox16.Text, out Cgv); double.TryParse(textBox62.Text, out C13); double.TryParse(textBox63.Text, out C30); double.TryParse(textBox64.Text, out C45);
            double.TryParse(textBox57.Text, out t1); double.TryParse(textBox39.Text, out tnv); double.TryParse(textBox58.Text, out kp); double.TryParse(textBox59.Text, out kt);
            double.TryParse(textBox60.Text, out tvv); double.TryParse(textBox61.Text, out tpodp);
            if (combobox9.Text == "неактивен")
            {
                double.TryParse(textBox1.Text, out Qt); double.TryParse(textBox2.Text, out Qpm13); double.TryParse(textBox3.Text, out Dp30); double.TryParse(textBox4.Text, out D140_45); double.TryParse(textBox13.Text, out Gsetv1); double.TryParse(textBox22.Text, out Gsetv3);
                double.TryParse(textBox19.Text, out Gsetv4); double.TryParse(textBox20.Text, out Gsetv6); double.TryParse(textBox15.Text, out tos1); double.TryParse(textBox18.Text, out tos3); double.TryParse(textBox21.Text, out tos4); double.TryParse(textBox40.Text, out tos6);
                double.TryParse(textBox14.Text, out Ghob); double.TryParse(textBox28.Text, out Gsw); double.TryParse(textBox29.Text, out Gmh); double.TryParse(textBox30.Text, out tsw); double.TryParse(textBox70.Text, out G13pr); double.TryParse(textBox5.Text, out Gsetw);
                double.TryParse(textBox7.Text, out Gdist); double.TryParse(textBox8.Text, out Gpts); double.TryParse(textBox6.Text, out Ns); double.TryParse(textBox10.Text, out Cn); double.TryParse(textBox9.Text, out Cnsn); double.TryParse(textBox11.Text, out Ct);
                double.TryParse(textBox12.Text, out bteh); double.TryParse(textBox15.Text, out tos1); double.TryParse(textBox18.Text, out tos3); double.TryParse(textBox21.Text, out tos4);
                double.TryParse(textBox40.Text, out tos6);
            }

            int t = 0; int tex = 1; int ot = 0; double MMM3 = -100000000000000;

            if (combobox9.Text == "активен")
            {
                string cmb2 = combobox11.Text;
                string cmb3 = combobox12.Text;
                int.TryParse(cmb2, out cmb4); int.TryParse(cmb3, out cmb5);
                ISHD = ishd(cmb4, cmb5);
                excelapp.Visible = true;
                excelappworkbooks = excelapp.Workbooks;
                //Открываем книгу и получаем на нее ссылку
                excelappworkbook = excelapp.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + @"\Результаты.xlsx",
                 Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing, Type.Missing);
                excelsheets = excelappworkbook.Worksheets;
                tex = cmb4;
            }
            else { cmb4 = 1; cmb5 = 1; }

            for (int j = cmb4; j <= cmb5; j++)
            {
                t = 0;
                for (int i1 = 0; i1 <= 1; i1++)
                {
                    for (int i2 = 0; i2 <= 1; i2++)
                    {
                        for (int i3 = 0; i3 <= 1; i3++)
                        {
                            for (int i4 = 0; i4 <= 1; i4++)
                            {
                                for (int i5 = 0; i5 <= 1; i5++)
                                {
                                    for (int i6 = 0; i6 <= 1; i6++)
                                    {
                                        for (int iptn1 = 0; iptn1 <= 1; iptn1++)
                                        {
                                            for (int iptn2 = 0; iptn2 <= 1; iptn2++)
                                            {
                                                if (combobox7.Text != "ремонт" & iptn2 == 1 & iptn1 == 0) { continue; }
                                                if (i5 != 1 & (iptn1 == 1 || iptn2 == 1)) { continue; }
                                                if (t == 110)
                                                {
                                                    t = t;
                                                }

                                                if (combobox1.Text == "ремонт") { indtg1 = 2; i1 = 2; } if (combobox2.Text == "ремонт") { indtg2 = 2; i2 = 2; }
                                                if (combobox3.Text == "ремонт") { indtg3 = 2; i3 = 2; } if (combobox4.Text == "ремонт") { indtg4 = 2; i4 = 2; }
                                                if (combobox5.Text == "ремонт") { indtg5 = 2; i5 = 2; } if (combobox6.Text == "ремонт") { indtg6 = 2; i6 = 2; }
                                                if (combobox7.Text == "ремонт") { indptn1 = 2; iptn1 = 2; } if (combobox8.Text == "ремонт") { indptn2 = 2; iptn2 = 2; }
                                                if (combobox10.Text == "неактивен") { i1 = 2; i2 = 2; i3 = 2; i4 = 2; i5 = 2; i6 = 2; iptn1 = 2; iptn2 = 2; }

                                                iii1 = ""; iii2 = ""; iii3 = ""; iii4 = ""; iii5 = ""; iii6 = ""; iiiptn1 = ""; iiiptn2 = ""; kkk1 = ""; kkk2 = ""; kkk3 = ""; kkk4 = ""; kkk5 = ""; kkk6 = ""; kkk7 = "";

                                                if (i1 == 1) { indtg1 = 1; iii1 = "1"; } if (i1 == 0 || i1 == 2) { indtg1 = 0; iii1 = ""; }
                                                if (i2 == 1) { indtg2 = 1; iii2 = "2"; } if (i2 == 0 || i2 == 2) { indtg2 = 0; iii2 = ""; }
                                                if (i3 == 1) { indtg3 = 1; iii3 = "3"; } if (i3 == 0 || i3 == 2) { indtg3 = 0; iii3 = ""; }
                                                if (i4 == 1) { indtg4 = 1; iii4 = "4"; } if (i4 == 0 || i4 == 2) { indtg4 = 0; iii4 = ""; }
                                                if (i5 == 1) { indtg5 = 1; iii5 = "5"; } if (i5 == 0 || i5 == 2) { indtg5 = 0; iii5 = ""; }
                                                if (i6 == 1) { indtg6 = 1; iii6 = "6"; } if (i6 == 0 || i6 == 2) { indtg6 = 0; iii6 = ""; }
                                                if (iptn1 == 1) { indptn1 = 1; iiiptn1 = "1"; } if (iptn1 == 0 || iptn1 == 2) { indptn1 = 0; iiiptn1 = ""; }
                                                if (iptn2 == 1) { indptn2 = 1; iiiptn2 = "2"; } if (iptn2 == 0 || iptn2 == 2) { indptn2 = 0; iiiptn2 = ""; }

                                                indk1 = 0; indk2 = 0; indk3 = 0; indk4 = 0; indk5 = 0; indk6 = 0;

                                                if (combobox10.Text == "неактивен")
                                                {
                                                    if (combobox1.Text == "работа") indtg1 = 1; if (combobox2.Text == "работа") indtg2 = 1; if (combobox3.Text == "работа") indtg3 = 1;
                                                    if (combobox4.Text == "работа") indtg4 = 1; if (combobox5.Text == "работа") indtg5 = 1; if (combobox6.Text == "работа") indtg6 = 1;
                                                    if (combobox7.Text == "работа") indptn1 = 1; if (combobox8.Text == "работа") indptn2 = 1;
                                                    if (combobox13.Text == "работа") indk1 = 1; if (combobox14.Text == "работа") indk2 = 1; if (combobox15.Text == "работа") indk3 = 1;
                                                    if (combobox16.Text == "работа") indk4 = 1; if (combobox17.Text == "работа") indk5 = 1; if (combobox18.Text == "работа") indk6 = 1;
                                                    if (combobox19.Text == "работа") indk7 = 1;
                                                    if (indtg1 == 0 & indtg2 == 0 & indtg3 == 0 & indtg4 == 0 & indtg5 == 0 & indtg6 == 0) { MessageBox.Show("Не выбраны работающие турбоагрегаты", "ОШИБКА", MessageBoxButton.OK, MessageBoxImage.Stop); tabPage1.Focus(); return; }
                                                }

                                                if (Gsetv1 != 0) indG1 = 1; if (Gsetv3 != 0) indG3 = 1; if (Gsetv4 != 0) indG4 = 1; if (Gsetv6 != 0) indG6 = 1;
                                                double Gsetwm = Gsetw;
                                                if (indtg6 == 1 & indG6 == 0) { Gsetv6 = 3400; if (Gsetwm < 3400) Gsetv6 = Gsetwm; Gsetwm = Gsetwm - Gsetv6; }
                                                if (indtg4 == 1 & indG4 == 0) { Gsetv4 = 5000; if (Gsetwm < 5000) Gsetv4 = Gsetwm; Gsetwm = Gsetwm - Gsetv4; }
                                                if (indtg3 == 1 & indG3 == 0) { Gsetv3 = 3500; if (Gsetwm < 3500) Gsetv3 = Gsetwm; Gsetwm = Gsetwm - Gsetv3; }
                                                if (indtg1 == 1 & indG1 == 0) { Gsetv1 = 2500; if (Gsetwm < 2500) Gsetv1 = Gsetwm; }
                                                
                                                Null();                                                   
                                                if (combobox9.Text == "активен")
                                                {
                                                    Ns = ISHD[0, j]; Qt = ISHD[1, j]; Qpm13 = ISHD[2, j]; Dp30 = ISHD[3, j]; D140_45 = ISHD[4, j]; Ghob = ISHD[5, j]; Gdist = ISHD[6, j]; Gsw = ISHD[7, j]; Gmh = ISHD[8, j]; G13pr = ISHD[9, j]; tsw = ISHD[10, j];
                                                    Gsetw = ISHD[11, j]; Gpts = ISHD[12, j]; Cn = ISHD[13, j]; Cnsn = ISHD[14, j]; Ct = ISHD[15, j]; bteh = ISHD[16, j]; tos1 = ISHD[17, j]; tos3 = ISHD[18, j];
                                                    tos4 = ISHD[19, j]; tos6 = ISHD[20, j]; Gsetv1 = ISHD[21, j]; Gsetv3 = ISHD[22, j]; Gsetv4 = ISHD[23, j]; Gsetv6 = ISHD[24, j];
                                                }

                                                if (indtg2 == 1) { Dp2max = 374; Dp2min = 130; } else { Dp2max = 0; Dp2min = 0; }
                                                if (indtg1 == 1) { N1max = 35; N1min = 5; Qt1max = 110; }
                                                else { N1max = 0; N1min = 0; Qt1max = 0; }
                                                if (indtg3 == 1) { N3max = 60; N3min = 30; Qt3max = 92; Qt31 = 60; }
                                                else { N3max = 0; N3min = 0; Qt3max = 0; Qt31 = 0; }
                                                if (indtg4 == 1) { N4max = 120; N4min = 40; Qt4max = 168; Qt41 = 100; }
                                                else { N4max = 0; N4min = 0; Qt4max = 0; Qt41 = 0; }
                                                if (indtg5 == 1) { Dp5max = 410; Dp5min = 105; D30_13max = 120; Q30_13min = 0; }
                                                else { D30_13max = 0; Q30_13min = 0; }
                                                if (indtg6 == 1) { N6max = 165; N6min = 40; Qt6max = 160; Dp6max = 460; Qt61 = 90; }
                                                else { N6max = 0; N6min = 0; Qt6max = 0; Qt61 = 0; }
                                                Qbmax = 150;
                                                if (indptn1 == 1 || indptn2 == 1) Dptnmax = 91; Dptnmin = 70;
                                                if ((indptn1 == 1 & ((indptn2 == 0 || indptn2 == 1))) || (indptn2 == 1 & ((indptn1 == 0 || indptn1 == 2)))) Dptnmax = 100;
                                                D140_13max = 500;
                                                if (Qt - Qpodp < Qt1max) Qt1max = Qt; if (Qt - Qpodp < Qt3max) Qt3max = Qt; if (Qt - Qpodp < Qt4max) Qt4max = Qt; if (Qt - Qpodp < Qt6max) Qt6max = Qt;
                                                if (Qt - Qpodp < Qbmax) Qbmax = Qt; if (Qt - Qpodp < Qpvk) Qpvkmax = Qt;

                                                if (indtg3 == 0) { indP3 = 0; }
                                                if (indtg4 == 0) { indP4 = 0; }
                                                if (indtg6 == 0) { indP6 = 0; }

                                                if (checkbox19.IsChecked == true & indtg1 == 1) double.TryParse(textBox17.Text, out N1); if (checkbox20.IsChecked == true & indtg1 == 1) double.TryParse(textBox23.Text, out Qt1);
                                                if (checkbox29.IsChecked == true & indtg2 == 1) double.TryParse(textBox24.Text, out Dp2);
                                                if (checkbox28.IsChecked == true & indtg3 == 1) double.TryParse(textBox25.Text, out N3); if (checkbox27.IsChecked == true & indtg3 == 1) double.TryParse(textBox26.Text, out Qt3);
                                                if (checkbox26.IsChecked == true & indtg4 == 1) double.TryParse(textBox27.Text, out N4); if (checkbox25.IsChecked == true & indtg4 == 1) double.TryParse(textBox31.Text, out Qt4);
                                                if (checkbox24.IsChecked == true & indtg5 == 1) double.TryParse(textBox32.Text, out Dp5);
                                                if (checkbox23.IsChecked == true & indtg6 == 1) double.TryParse(textBox33.Text, out N6); if (checkbox22.IsChecked == true & indtg6 == 1) double.TryParse(textBox34.Text, out Qt6); if (checkbox21.IsChecked == true & indtg6 == 1) double.TryParse(textBox35.Text, out Dp6);
                                                if (checkbox30.IsChecked == true) double.TryParse(textBox36.Text, out D30_13);
                                                if (checkbox7.IsChecked == true & indptn1 == 1) double.TryParse(textBox37.Text, out Dptn1); if (checkbox8.IsChecked == true & indptn2 == 1) double.TryParse(textBox38.Text, out Dptn2);
                                                if (checkbox5.IsChecked == true) double.TryParse(textBox43.Text, out D140_13); if (checkbox6.IsChecked == true) double.TryParse(textBox44.Text, out D140_30);
                                                if (checkbox1.IsChecked == true) double.TryParse(textBox45.Text, out dtk1); if (checkbox2.IsChecked == true) double.TryParse(textBox46.Text, out dtk3);
                                                if (checkbox3.IsChecked == true) double.TryParse(textBox47.Text, out dtk4); if (checkbox4.IsChecked == true) double.TryParse(textBox48.Text, out dtk6);
                                                double.TryParse(textBox49.Text, out dtsp1); double.TryParse(textBox50.Text, out dtsp3); double.TryParse(textBox51.Text, out dtsp4); double.TryParse(textBox52.Text, out dtsp6);
                                                if (checkbox35.IsChecked == true) double.TryParse(textBox53.Text, out P1); if (checkbox36.IsChecked == true) double.TryParse(textBox54.Text, out P3);
                                                if (checkbox37.IsChecked == true) double.TryParse(textBox55.Text, out P4); if (checkbox38.IsChecked == true) double.TryParse(textBox56.Text, out P6);

                                                if (indtg1 == 1 & tos1 == 0) { MessageBox.Show("Не задана температура обратной сетевой воды ТГ-1", "ОШИБКА", MessageBoxButton.OK, MessageBoxImage.Stop); tabPage1.Focus(); return; }
                                                if (indtg3 == 1 & tos3 == 0) { MessageBox.Show("Не задана температура обратной сетевой воды ТГ-3", "ОШИБКА", MessageBoxButton.OK, MessageBoxImage.Stop); tabPage1.Focus(); return; }
                                                if (indtg4 == 1 & tos4 == 0) { MessageBox.Show("Не задана температура обратной сетевой воды ТГ-4", "ОШИБКА", MessageBoxButton.OK, MessageBoxImage.Stop); tabPage1.Focus(); return; }
                                                if (indtg6 == 1 & tos6 == 0) { MessageBox.Show("Не задана температура обратной сетевой воды ТГ-6", "ОШИБКА", MessageBoxButton.OK, MessageBoxImage.Stop); tabPage1.Focus(); return; }

                                                X2 = but();

                                                if (X2[13] > 1 & X2[8] < 1 & X2[40] < 1) { MessageBox.Show("Для работы РОУ-30/13 недостаточно источников пара 30 ата", "ОШИБКА", MessageBoxButton.OK, MessageBoxImage.Stop); tabPage1.Focus(); return; }

                                                for (int p = 0; p <= 99; p++)
                                                {
                                                    X3[t, p] = X2[p];
                                                }

                                                if (indk1 == 1) kkk1 = "1"; else kkk1 = "";
                                                if (indk2 == 1) kkk2 = "2"; else kkk2 = "";
                                                if (indk3 == 1) kkk3 = "3"; else kkk3 = "";
                                                if (indk4 == 1) kkk4 = "4"; else kkk4 = "";
                                                if (indk5 == 1) kkk5 = "5"; else kkk5 = "";
                                                if (indk6 == 1) kkk6 = "6"; else kkk6 = "";
                                                if (indk7 == 1) kkk7 = "7"; else kkk7 = "";
                                                if (X3[t, 47] > 0) iiiptn1 = "1"; else iiiptn1 = "";
                                                if (X3[t, 48] > 0) iiiptn2 = "2"; else iiiptn2 = "";

                                                if (combobox10.Text == "неактивен")
                                                {
                                                    if (combobox20.Text == "разрешено")
                                                    {
                                                        if (indk1 == 1) combobox13.Text = "работа"; if (indk2 == 1) combobox14.Text = "работа"; if (indk3 == 1) combobox15.Text = "работа";
                                                        if (indk4 == 1) combobox16.Text = "работа"; if (indk5 == 1) combobox17.Text = "работа"; if (indk6 == 1) combobox18.Text = "работа";
                                                        if (indk7 == 1) combobox19.Text = "работа";
                                                    }
                                                }

                                                if (combobox10.Text == "активен")
                                                {
                                                    combobox21.Items.Add((t + 1).ToString() + "  " + "ТГ:" + " " + iii1 + iii2 + iii3 + iii4 + iii5 + iii6 + " " + "ПТН:" + " " + iiiptn1 + iiiptn2);
                                                    if (iii1 == "1") X3[t, 50] = 1; if (iii2 == "2") X3[t, 51] = 1; if (iii3 == "3") X3[t, 52] = 1;
                                                    if (iii4 == "4") X3[t, 53] = 1; if (iii5 == "5") X3[t, 54] = 1; if (iii6 == "6") X3[t, 55] = 1;
                                                    if (kkk1 == "1") X3[t, 56] = 1; if (kkk2 == "2") X3[t, 57] = 1; if (kkk3 == "3") X3[t, 58] = 1;
                                                    if (kkk4 == "4") X3[t, 59] = 1; if (kkk5 == "5") X3[t, 60] = 1; if (kkk6 == "6") X3[t, 61] = 1;
                                                    if (kkk7 == "7") X3[t, 62] = 1;
                                                    if (X3[t, 38] > MMM3 & X3[t, 38] != 0)
                                                    {
                                                        MMM3 = X3[t, 38]; ot = t;
                                                        iii11 = iii1; iii21 = iii2; iii31 = iii3; iii41 = iii4; iii51 = iii5; iii61 = iii6;
                                                        kkk11 = kkk1; kkk21 = kkk2; kkk31 = kkk3; kkk41 = kkk4; kkk51 = kkk5; kkk61 = kkk6; kkk71 = kkk7;
                                                        iiiptn11 = iiiptn1; iiiptn21 = iiiptn2;
                                                    }
                                                }
                                                if (combobox10.Text == "активен") t++;                                                                                               
                                            }                                            
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (combobox9.Text == "активен")
                {
                    input(X3, ot, tex);
                }
                GC.Collect(); // убрать за собой
                if (combobox9.Text == "активен") tex++;
            }

            if (combobox10.Text == "активен")
            {
                combobox21.Text = (ot + 1).ToString() + "  " + "ТГ:" + " " + iii11 + iii21 + iii31 + iii41 + iii51 + iii61 + " " + "ПТН:" + " " + iiiptn11 + iiiptn21;
                textBox41.Text = iii11 + iii21 + iii31 + iii41 + iii51 + iii61 + " " + "ПТН:" + " " + iiiptn11 + iiiptn21;
                textBox42.Text = kkk11 + kkk21 + kkk31 + kkk41 + kkk51 + kkk61;
                output(X3, ot);
            }

            if (combobox9.Text == "неактивен" & combobox10.Text == "неактивен") output(X3, t);
            if (combobox9.Text == "активен" & combobox10.Text == "неактивен")
            {
                output(X3, cmb4);
                //Сохраняем результат
                excelappworkbooks = excelapp.Workbooks;
                excelappworkbook = excelappworkbooks[1];
                try { excelappworkbook.Save(); }
                catch (Exception) { excelappworkbook.Close(true, Type.Missing, Type.Missing); }
            }
            sWatch.Stop();
            if (combobox9.Text == "неактивен") tabPage1.Focus();
        }

        public double[] but()
        {
            if (indtg3 == 0 || indtg3 == 2) { indP3 = 0; }
            if (indtg4 == 0 || indtg4 == 2) { indP4 = 0; }
            if (indtg6 == 0 || indtg6 == 2) { indP6 = 0; }

            indcn1 = 0; indcn3 = 0; indcn4 = 0;
            if (indtg6 == 1) Gcw6 = 9000;
            if (indtg4 == 1) { if (radioButton5.IsChecked == true) indcn4 = 1; if (radioButton6.IsChecked == true) indcn4 = 2; }
            if (indtg3 == 1) { if (radioButton4.IsChecked == true) indcn3 = 1; if (radioButton3.IsChecked == true) indcn3 = 2; }
            if (indtg1 == 1) { if (radioButton11.IsChecked == true) indcn1 = 1; if (radioButton14.IsChecked == true) indcn1 = 2; }

            kpb = 0.55; hdop = 733.2; h1 = 0.5342 * t1 + 566.66; hv = 165; h13 = 695; h30 = 703.2; h45 = 725; hpp = 832.9;
            k13 = (h13 - tvv) / 1000; k30 = (h30 - tvv) / 1000; k45 = (h45 - tvv) / 1000;

            Qs = Qt + Qpm13 * k13 + Dp30 * k30 + D140_45 * k45; Qpodp = Gpts * (tpodp - tvv) / 1000;

            a = 1;

            Y = Dinamyc();
            Dp2 = Y[9]; Dp5 = Y[10]; D140_30 = Y[11]; D30_13 = Y[12]; D140_13 = Y[13]; Dptn1 = Y[14]; Dptn2 = Y[15]; P1 = Y[31]; P3 = Y[32]; P4 = Y[33]; P6 = Y[34];
            Dp6 = Y[0]; Qt6 = Y[1]; N6 = Y[2]; Qt4 = Y[3]; N4 = Y[4]; Qt3 = Y[5]; N3 = Y[6]; N1 = Y[7]; Qt1 = Y[8]; Qpb = Y[28]; nk = Y[30];

            indcn1 = Convert.ToInt16(Y[35]); indcn3 = Convert.ToInt16(Y[36]); indcn4 = Convert.ToInt16(Y[37]); indP3 = Convert.ToInt16(Y[38]);
            indP4 = Convert.ToInt16(Y[39]);
            
            Gek = (Q2(Dp2) + Q3(P3, Qt3, N3) + Q4(P4, Qt4, N4) + Q5(Dp5) + Q6(Dp6, Qt6, N6)) / 0.595 + D140_13 + D140_30 + D140_45;
            im = 1;
            nk = nek(Gek, im, out G1p, out G2p, out G3p, out G4p, out G5p, out G6p, out G7p);
            im = 0;
            B = Fin(out S, out Nsn, out Kotrt, out Kotrk, out nnet, out ntp, out q);

            if (bteh == 0) { be = (q * Kotrt) / (7 * Kotrk * nnet * ntp / 10000); bte = (B - (be * (S - Nko - Nto - Nsnsw(Gsetw))) / 1000) / Qs * 1000; }
            else { bte = bteh; be = (B - bte * Qs / 1000) / (S - Nsn) * 1000; }

            if (nk == 0 || nk == 100)
            {
                for (int i = 0; i < 100; i++)
                {
                    VAR[i] = 0;
                }
            }
            else
            {
                VAR[0] = N1; VAR[1] = Qt1; VAR[2] = NN2(Dp2); VAR[3] = Dp2; VAR[4] = N3; VAR[5] = Qt3; VAR[6] = N4; VAR[7] = Qt4; VAR[8] = Dp5; VAR[9] = Dp6; VAR[10] = Qt6; VAR[11] = N6;
                VAR[12] = Qpb / kpb; VAR[13] = D30_13; VAR[14] = G1p; VAR[15] = G2p; VAR[16] = G3p; VAR[17] = G4p; VAR[18] = G5p; VAR[19] = G6p; VAR[20] = G7p;
                VAR[21] = Q2(Dp2) / 0.595; VAR[22] = Q3(P3, Qt3, N3) / 0.595; VAR[23] = Q4(P4, Qt4, N4) / 0.595; VAR[24] = Q5(Dp5) / 0.595; VAR[25] = Q6(Dp6, Qt6, N6) / 0.595;
                VAR[26] = D140_45; VAR[27] = Ghob * 64 / 536; VAR[28] = Y[50]; VAR[29] = Dpsv; VAR[30] = be; VAR[31] = bte; VAR[32] = Nsn; VAR[33] = q; VAR[34] = nk;
                VAR[35] = Daux(); VAR[36] = B; VAR[37] = Qpvk; VAR[39] = D140_13; VAR[40] = D140_30; VAR[41] = Qpb; VAR[42] = S;
                VAR[44] = Dp30; VAR[45] = Y[49]; VAR[46] = Qpm13; VAR[47] = Dptn1; VAR[48] = Dptn2; VAR[49] = Y[41]; VAR[63] = Y[35]; VAR[64] = Y[36]; VAR[65] = Y[37]; VAR[66] = Y[38]; VAR[67] = Y[39]; VAR[68] = Y[40];
                VAR[69] = Y[42]; VAR[70] = Y[43]; VAR[71] = Y[44]; VAR[72] = Y[45]; VAR[38] = Y[48]; VAR[73] = Y[46] / 1000; VAR[74] = P1; VAR[75] = P3; VAR[76] = P4; VAR[77] = P6;
            }
            return VAR;
        }

        public double[] Dinamyc()
        {
            int chh = 0;
            ch3 = 0; MMM2 = -1000000000000;

            if (radioButton1.IsChecked == true) indP3 = 1; if (radioButton2.IsChecked == true) indP3 = 2;
            if (radioButton8.IsChecked == true) indP4 = 1; if (radioButton10.IsChecked == true) indP4 = 2;
            if (radioButton13.IsChecked == true) indP6 = 1; if (radioButton12.IsChecked == true) indP6 = 2;

            for (tp1 = 1; tp1 < 4; tp1++)
            {
                if (indtg1 != 1 || checkbox20.IsChecked == true) tp1 = 4;
                for (np1 = 1; np1 < 5; np1++)
                {
                    if (indtg1 != 1 || checkbox19.IsChecked == true) np1 = 5;
                    for (tp3 = 1; tp3 < 4; tp3++)
                    {
                        if (indtg3 != 1 || checkbox27.IsChecked == true) tp3 = 4;
                        for (np3 = 1; np3 < 5; np3++)
                        {
                            if (indtg3 != 1 || checkbox28.IsChecked == true) np3 = 5;
                            for (tp4 = 1; tp4 < 4; tp4++)
                            {
                                if (indtg4 != 1 || checkbox25.IsChecked == true) tp4 = 4;
                                for (np4 = 1; np4 < 5; np4++)
                                {
                                    if (indtg4 != 1 || checkbox26.IsChecked == true) np4 = 5;
                                    for (tp6 = 1; tp6 < 4; tp6++)
                                    {
                                        if (indtg6 != 1 || checkbox22.IsChecked == true) tp6 = 4;
                                        for (np6 = 1; np6 < 4; np6++)
                                        {
                                            if (indtg6 != 1 || checkbox23.IsChecked == true) np6 = 4;
                                            for (dptn1 = 2; dptn1 < 3; dptn1++)
                                            {
                                                if (indptn1 == 0 || indptn1 == 2 || checkbox7.IsChecked == true) dptn1 = 3;
                                                for (dptn2 = 2; dptn2 < 3; dptn2++)
                                                {
                                                    if (indptn2 == 0 || indptn2 == 2 || checkbox8.IsChecked == true) dptn2 = 3;
                                                    for (dp5 = 1; dp5 < 3; dp5++)
                                                    {
                                                        if (indtg5 != 1 || checkbox24.IsChecked == true) dp5 = 3;
                                                        for (dp6 = 1; dp6 < 4; dp6++)
                                                        {
                                                            for (int nvar = 1; nvar < 7; nvar++)
                                                            {
                                                                if (indtg6 != 1 || checkbox21.IsChecked == true) dp6 = 4;
                                                                for (dp2 = 1; dp2 < 3; dp2++)
                                                                {
                                                                    ch3++;
                                                                    if (ch3 == 98)
                                                                    {
                                                                        ch3 = ch3;
                                                                    }

                                                                    if (indtg1 == 1 & checkbox19.IsChecked == true & checkbox20.IsChecked == true)
                                                                    {
                                                                        if (indG1 == 0) { Gsetv1 = Gsetw * Qt1 / (Qt1 + Qt3 + Qt4 + Qt6 + 1); if (Gsetv1 > 2500) Gsetv1 = 2500; }
                                                                        P1 = Pn(Qt1 * 1000 / (Gsetv1 + 1) + tos1 + dtsp1); if (P1 > 2.5) P1 = 2.5; if (P1 < 0.3) P1 = 0.3;
                                                                        if (((Qt1 < 31.674 * Math.Pow(P1, 2) - 163.35 * P1 + 210.41 & N1 > ogr_TG1rb(P1, Qt1)) || (Qt1 > 31.674 * Math.Pow(P1, 2) - 163.35 * P1 + 210.41 & N1 > ogr_TG1rn(P1, Qt1))) || N1 < ogr_TG1l(P1, Qt1)) continue;
                                                                    }
                                                                    if (checkbox24.IsChecked == true & (Dp2 < Dp2min || checkbox24.IsChecked == true & Dp2 > Dp2max)) continue;
                                                                    if (indtg3 == 1 & checkbox28.IsChecked == true & checkbox27.IsChecked == true)
                                                                    {
                                                                        if (indG3 == 0) { Gsetv3 = Gsetw * Qt3 / (Qt1 + Qt3 + Qt4 + Qt6 + 1); if (Gsetv3 > 3500) Gsetv3 = 3500; }
                                                                        P3 = Pn(Qt3 * 1000 / (Gsetv3 + 1) + tos3 + dtsp3); if (P3 > 2.5) P3 = 2.5; if (P3 < 0.3) P3 = 0.3;
                                                                        if (N3 > ogr_TG3r(P3, Qt3) || N3 < ogr_TG3l(P3, Qt3)) continue;
                                                                    }
                                                                    if (indtg4 == 1 & checkbox26.IsChecked == true & checkbox25.IsChecked == true)
                                                                    {
                                                                        if (indG4 == 0) { Gsetv4 = Gsetw * Qt4 / (Qt1 + Qt3 + Qt4 + Qt6 + 1); if (Gsetv4 > 5000) Gsetv4 = 5000; }
                                                                        P4 = Pn(Qt4 * 1000 / (Gsetv4 + 1) + tos4 + dtsp4); if (P4 > 2.5) P4 = 2.5; if (P4 < 0.3) P4 = 0.3;
                                                                        if ((N4 > ogr_TG4r(P4, Qt4) || N4 > 120) || N4 < ogr_TG4l(P4, Qt4)) continue;
                                                                    }
                                                                    if (checkbox24.IsChecked == true & (Dp5 > Dp5max || Dp5 < Dp5min)) continue;

                                                                    if ((checkbox22.IsChecked == false & checkbox23.IsChecked == false) || (checkbox22.IsChecked == false & checkbox24.IsChecked == false) || (checkbox23.IsChecked == false & checkbox24.IsChecked == false))
                                                                    {
                                                                        if ((np6 == 3 || checkbox23.IsChecked == true) & (nvar == 2 || nvar == 3 || nvar == 5 || nvar == 6)) continue;
                                                                        if (checkbox22.IsChecked == true & (nvar == 1 || nvar == 2 || nvar == 4 || nvar == 5)) continue;
                                                                        if (checkbox24.IsChecked == true & (nvar == 1 || nvar == 3 || nvar == 4 || nvar == 6)) continue;
                                                                    }
                                                                    if (checkBox10.IsChecked == false & (np1 == 3 || np3 == 3 || np4 == 3 || np6 == 3)) continue;

                                                                    if (np1 == 3 && (np3 == 3 || np4 == 3 || np6 == 3)) continue;
                                                                    if (np3 == 3 && (np4 == 3 || np6 == 3)) continue;
                                                                    if (np4 == 3 && np6 == 3) continue;
                                                                    if (tp1 == 3 && (tp3 == 3 || tp4 == 3 || tp6 == 3)) continue;
                                                                    if (tp3 == 3 && (tp4 == 3 || tp6 == 3)) continue;
                                                                    if (tp4 == 3 && tp6 == 3) continue;

                                                                    if (checkbox19.IsChecked == false) N1 = 0; if (checkbox29.IsChecked == false) N2 = 0; if (checkbox28.IsChecked == false) N3 = 0;
                                                                    if (checkbox26.IsChecked == false) N4 = 0; if (checkbox24.IsChecked == false) N5 = 0; if (checkbox23.IsChecked == false) N6 = 0;
                                                                    if (checkbox7.IsChecked == false) Dptn1 = 0; if (checkbox8.IsChecked == false) Dptn2 = 0; if (checkbox22.IsChecked == false) Qt6 = 0; if (checkbox25.IsChecked == false) Qt4 = 0;
                                                                    if (checkbox27.IsChecked == false) Qt3 = 0; if (checkbox20.IsChecked == false) Qt1 = 0; if (checkbox5.IsChecked == false) D140_13 = 0; if (checkbox6.IsChecked == false) D140_30 = 0;
                                                                    if (checkbox30.IsChecked == false) D30_13 = 0; if (checkbox21.IsChecked == false) Dp6 = 0; if (checkbox29.IsChecked == false) Dp2 = 0; if (checkbox24.IsChecked == false) Dp5 = 0; Qpb = 0;
                                                                    if (radioButton1.IsChecked == false & radioButton2.IsChecked == false) indP3 = 0; if (radioButton8.IsChecked == false & radioButton10.IsChecked == false) indP4 = 0;
                                                                    if (radioButton13.IsChecked == false & radioButton12.IsChecked == false) indP6 = 0;
                                                                    if (radioButton11.IsChecked == false & radioButton14.IsChecked == false) indcn1 = 0; if (radioButton4.IsChecked == false & radioButton3.IsChecked == false) indcn3 = 0;
                                                                    if (radioButton5.IsChecked == false & radioButton6.IsChecked == false) indcn4 = 0;
                                                                    
                                                                    if (indtg2 != 1 || checkbox29.IsChecked == true) dp2 = 3;

                                                                    if (tp1 == 1) Qt1 = 0;
                                                                    if (tp1 == 2) Qt1 = Qt1max;
                                                                    if (tp3 == 1) Qt3 = 0;
                                                                    if (tp3 == 2) Qt3 = Qt3max;
                                                                    if (tp4 == 1) Qt4 = 0;
                                                                    if (tp4 == 2) Qt4 = Qt4max;
                                                                    if (tp6 == 1) Qt6 = 0;
                                                                    if (tp6 == 2) Qt6 = Qt6max;

                                                                    if (Qt - Qpodp < Qt1 + Qt3 + Qt4 + Qt6) continue;

                                                                    if (tp1 == 3) Qt1 = Qt - Qpodp- Qt3 - Qt4 - Qt6;
                                                                    if (tp3 == 3) Qt3 = Qt - Qpodp- Qt1 - Qt4 - Qt6;
                                                                    if (tp4 == 3) Qt4 = Qt - Qpodp- Qt1 - Qt3 - Qt6;
                                                                    if (tp6 == 3) Qt6 = Qt - Qpodp- Qt1 - Qt3 - Qt4;

                                                                    if (Qt1 > Qt1max || Qt3 > Qt3max || Qt4 > Qt4max || Qt6 > Qt6max) continue;

                                                                    if (indG1 == 0) { Gsetv1 = Gsetw * Qt1 / (Qt1 + Qt3 + Qt4 + Qt6 + 1); if (Gsetv1 > 2500) Gsetv1 = 2500; }
                                                                    if (indG3 == 0) { Gsetv3 = Gsetw * Qt3 / (Qt1 + Qt3 + Qt4 + Qt6 + 1); if (Gsetv3 > 3500) Gsetv3 = 3500; }
                                                                    if (indG4 == 0) { Gsetv4 = Gsetw * Qt4 / (Qt1 + Qt3 + Qt4 + Qt6 + 1); if (Gsetv4 > 5000) Gsetv4 = 5000; }
                                                                    if (indG6 == 0) { Gsetv6 = Gsetw * Qt6 / (Qt1 + Qt3 + Qt4 + Qt6 + 1); if (Gsetv6 > 3400) Gsetv6 = 3400; }

                                                                    if (indtg1 == 1) { P1 = Pn(Qt1 * 1000 / (Gsetv1 + 1) + tos1 + dtsp1); if (P1 > 2.5) P1 = 2.5; if (P1 < 0.3) P1 = 0.3; }
                                                                    if (indtg3 == 1) { P3 = Pn(Qt3 * 1000 / (Gsetv3 + 1) + tos3 + dtsp3); if (P3 > 2.5) P3 = 2.5; if (P3 < 0.3) P3 = 0.3; }
                                                                    if (indtg4 == 1) { P4 = Pn(Qt4 * 1000 / (Gsetv4 + 1) + tos4 + dtsp4); if (P4 > 2.5) P4 = 2.5; if (P4 < 0.3) P4 = 0.3; }
                                                                    if (indtg6 == 1) { P6 = Pn(Qt6 * 1000 / (Gsetv6 + 1) + tos6 + dtsp6); if (P6 > 2.5) P6 = 2.5; if (P6 < 0.3) P6 = 0.3; }

                                                                    if (Qt3 < Qt31 & indP3 == 0) indP3 = 1; if (Qt3 > Qt31 & indP3 == 0) indP3 = 2;
                                                                    if (Qt4 < Qt41 & indP4 == 0) indP4 = 1; if (Qt4 > Qt41 & indP4 == 0) indP4 = 2;
                                                                    if (Qt6 < Qt61 & indP6 == 0) indP6 = 1; if (Qt6 > Qt61 & indP6 == 0) indP6 = 2;

                                                                    if ((indtg1 == 1 & checkbox19.IsChecked == true & checkbox20.IsChecked == false))
                                                                    {
                                                                        Nv1 = ogrTG1n(P1);
                                                                        if (N1 < Nv1 && Qt1 > ogr_TG1Qtl(P1, N1)) { Qt1 = ogr_TG1Qtl(P1, N1); P1 = Pn(Qt1 * 1000 / (Gsetv1 + 1) + tos1 + dtsp1); if (P1 > 2.5) P1 = 2.5; if (P1 < 0.3) P1 = 0.3; }
                                                                        if (N1 > Nv1 && Qt1 > ogr_TG1Qtrn(P1, N1)) { Qt1 = ogr_TG1Qtrn(P1, N1); P1 = Pn(Qt1 * 1000 / (Gsetv1 + 1) + tos1 + dtsp1); if (P1 > 2.5) P1 = 2.5; if (P1 < 0.3) P1 = 0.3; }
                                                                        if (Qt1 < ogr_TG1Qtrb(P1, N1)) { Qt1 = ogr_TG1Qtrb(P1, N1); P1 = Pn(Qt1 * 1000 / (Gsetv1 + 1) + tos1 + dtsp1); if (P1 > 2.5) P1 = 2.5; if (P1 < 0.3) P1 = 0.3; }
                                                                    }
                                                                    if ((indtg3 == 1 & checkbox28.IsChecked == true & checkbox27.IsChecked == false))
                                                                    {
                                                                        Nv3 = ogrTG3n(P3, Qt3);
                                                                        if (N3 > Nv3 & Qt3 > ogr_TG3Qtr(P3, N3)) { Qt3 = ogr_TG3Qtr(P3, N3); P3 = Pn(Qt3 * 1000 / (Gsetv3 + 1) + tos3 + dtsp3); if (P3 > 2.5) P3 = 2.5; if (P3 < 0.3) P3 = 0.3; }
                                                                        if (N3 < Nv3 & Qt3 > ogr_TG3Qtl(P3, N3)) { Qt3 = ogr_TG3Qtl(P3, N3); P3 = Pn(Qt3 * 1000 / (Gsetv3 + 1) + tos3 + dtsp3); if (P3 > 2.5) P3 = 2.5; if (P3 < 0.3) P3 = 0.3; }
                                                                    }
                                                                    if ((indtg4 == 1 & checkbox26.IsChecked == true & checkbox25.IsChecked == false))
                                                                    {
                                                                        Nv4 = ogrTG4n(P4, Qt4);
                                                                        if (N4 > Nv4 & Qt4 > ogr_TG4Qtr(P4, N4)) { Qt4 = ogr_TG4Qtr(P4, N4); P4 = Pn(Qt4 * 1000 / (Gsetv4 + 1) + tos4 + dtsp4); if (P4 > 2.5) P4 = 2.5; if (P4 < 0.3) P4 = 0.3; }
                                                                        if (N4 < Nv4 & Qt4 > ogr_TG4Qtl(P4, N4)) { Qt4 = ogr_TG4Qtl(P4, N4); }
                                                                    }

                                                                    Qpb = Qt - Qpodp- Qt1 - Qt3 - Qt4 - Qt6;
                                                                    if (Qpb < 0) continue;

                                                                    if (indtg1 == 1 & Gsetv1 != 0) P1 = Pn(Qt1 * 1000 / Gsetv1 + tos1 + dtsp1); else P1 = 0.5; if (P1 > 2.5) P1 = 2.5;
                                                                    if (indtg3 == 1 & Gsetv3 != 0) P3 = Pn(Qt3 * 1000 / Gsetv3 + tos3 + dtsp3); else P3 = 0.5; if (P3 > 2.5) P3 = 2.5;
                                                                    if (indtg4 == 1 & Gsetv4 != 0) P4 = Pn(Qt4 * 1000 / Gsetv4 + tos4 + dtsp4); else P4 = 0.5; if (P4 > 2.5) P4 = 2.5;
                                                                    if (indtg6 == 1 & Gsetv6 != 0) P6 = Pn(Qt6 * 1000 / Gsetv6 + tos6 + dtsp6); else P6 = 0.5; if (P6 > 2.5) P6 = 2.5;

                                                                    if (dptn1 == 2 & checkbox7.IsChecked == false) Dptn1 = 90;
                                                                    if (dptn2 == 2 & checkbox8.IsChecked == false) Dptn2 = 90;

                                                                    MMM = 0;

                                                                    if (indtg5 == 1) h30n = 772; else h30n = 772;
                                                                    if (checkbox19.IsChecked == true)
                                                                    {
                                                                        D1t = D1(P1, Qt1, N1);
                                                                        D1t = D1(P1, Qt1, N1 - PoprTG1t(t1, D1t));
                                                                        if (D1t < 0) D1t = 0;
                                                                    }
                                                                    double D13, D131;
                                                                    D13 = 100;
                                                                    int ch4 = 0;
                                                                    do
                                                                    {
                                                                        in6 = 0;
                                                                        D131 = D13;
                                                                        if (np1 == 1)
                                                                        {
                                                                            N1 = ogr_TG1l(P1, Qt1);
                                                                            if (N1 < N1min) N1 = N1min;
                                                                            D1t = D1(P1, Qt1, N1);
                                                                            if (D1t < 0) D1t = 0;
                                                                            N1 = N1 + PoprTG1t(t1, D1t);
                                                                        }                                                                        
                                                                        if (np1 == 2)
                                                                        {
                                                                            if (Qt1 < 31.674 * Math.Pow(P1, 2) - 163.35 * P1 + 210.41) N1 = ogr_TG1rb(P1, Qt1);
                                                                            else N1 = ogr_TG1rn(P1, Qt1);
                                                                            if (N1 < N1min) N1 = N1min;
                                                                            D1t = D1(P1, Qt1, N1);
                                                                            if (D1t < 0) D1t = 0;
                                                                            N1 = N1 + PoprTG1t(t1, D1t);                                                                            
                                                                        }
                                                                        if (np1 == 4)
                                                                        {
                                                                            N1 = N1min;
                                                                            Qt1 = ogr_TG1Qtl(P1, N1); P1 = Pn(Qt1 * 1000 / (Gsetv1 + 1) + tos1 + dtsp1); if (P1 > 2.5) P1 = 2.5; if (P1 < 0.3) P1 = 0.3;
                                                                            D1t = D1(P1, Qt1, N1);
                                                                            D1t = D1(P1, Qt1, N1 - PoprTG1t(t1, D1t));
                                                                        }
                                                                        if (np3 == 4)
                                                                        {
                                                                            N3 = N3min;
                                                                            Qt3 = ogr_TG3Qtl(P3, N3); P3 = Pn(Qt3 * 1000 / (Gsetv3 + 1) + tos3 + dtsp3); if (P3 > 2.5) P3 = 2.5; if (P3 < 0.3) P3 = 0.3;  
                                                                        }
                                                                        if (np4 == 4)
                                                                        {
                                                                            N4 = N4min;
                                                                            Qt4 = ogr_TG4Qtl(P4, N4); P4 = Pn(Qt4 * 1000 / (Gsetv4 + 1) + tos4 + dtsp4); if (P4 > 2.5) P4 = 2.5; if (P4 < 0.3) P4 = 0.3;
                                                                        }
                                                                        if (Qpb < 0) continue;
                                                                        if (np3 == 1) N3 = ogr_TG3l(P3, Qt3); if (N3 < N3min) N3 = N3min;
                                                                        if (np3 == 2) N3 = ogr_TG3r(P3, Qt3);
                                                                       
                                                                        if (np4 == 1) N4 = ogr_TG4l(P4, Qt4); if (N4 < N4min) N4 = N4min;
                                                                        if (np4 == 2) { N4 = ogr_TG4r(P4, Qt4); if (N4 > 120) N4 = 120; }

                                                                        if (np6 == 1) N6 = Kogr_TG6l(Dp6, Qt6);
                                                                        if (np6 == 2) N6 = ogr_TG6r(Dp6, Qt6);
                                                                        if (N6 < N6min) N6 = N6min;

                                                                        if (checkBox10.IsChecked == true && Ns < N1 + NN2(Dp2) + N3 + N4 + NN5(Dp5) + N6) { in6 = 1; break; }

                                                                        if (indtg1 == 1)
                                                                        {
                                                                            D3013 = (D1t * (h1 - h13) - (Dhptn(Dptn1) + Dhptn(Dptn2)) + (Dptn1 + Dptn2) * h13) / (h30n - h13);
                                                                            if (D3013 + Dptn1 + Dptn2 > D1t)
                                                                            {
                                                                                D3013 = ((Dptn1 + Dptn2) * h1 - Dhptn(Dptn1) - Dhptn(Dptn2)) / (h30n - h13);
                                                                            }
                                                                        }
                                                                        rty = (Dptn1 + Dptn2 + D3013 - D1t) * (h1 - hv) / (h13 - hv);

                                                                        if (dp2 == 1 || dp2 == 3)
                                                                        {
                                                                            if (dp2 == 1) Dp2 = Dp2min;
                                                                            if (dp6 != 1 & dp6 != 4)
                                                                            {
                                                                                rtn = Qpm13 + Qpb / kpb + D13 - rty - D2p(Dp2) - D30_13 * (h30 - hv) / (h13 - hv) - D140_13 * (hpp - hv) / (h13 - hv);
                                                                                if (rtn > Dp6max) rtn = Dp6max;
                                                                                Dp6 = D6(rtn, Qt6, N6);
                                                                                if (Dp6 < -2)
                                                                                {
                                                                                    break;
                                                                                }
                                                                                in6 = ogrTG6var(nvar);
                                                                                if (in6 == 1) break;
                                                                            }
                                                                        }
                                                                        if (dp2 == 2)
                                                                        {
                                                                            if (dp6 != 3)
                                                                            {
                                                                                rtn = Qpm13 + Qpb / kpb + D13 - rty - D6p(Dp6) - D30_13 * (h30 - hv) / (h13 - hv) - D140_13 * (hpp - hv) / (h13 - hv);
                                                                                Dp2 = DP2(rtn);
                                                                                if (Dp2 > Dp2max) Dp2 = Dp2max; if (Dp2 < Dp2min) Dp2 = Dp2min;
                                                                            }

                                                                            if (dp6 == 2)
                                                                            {
                                                                                rtn = Qpm13 + Qpb / kpb + D13 - rty - D2p(Dp2) - D30_13 * (h30 - hv) / (h13 - hv) - D140_13 * (hpp - hv) / (h13 - hv);
                                                                                if (rtn > Dp6max) rtn = Dp6max;
                                                                                Dp6 = D6(rtn, Qt6, N6);
                                                                                if (Dp6 < -2)
                                                                                {
                                                                                    break;
                                                                                }
                                                                                in6 = ogrTG6var(nvar);
                                                                                if (in6 == 1) break;
                                                                                Qpb = Qt - Qpodp- Qt1 - Qt3 - Qt4 - Qt6;
                                                                            }

                                                                            if (dp6 == 3)
                                                                            {
                                                                                rtn = Qpm13 + Qpb / kpb + D13 - rty - D2p(Dp2) - D30_13 * (h30 - hv) / (h13 - hv) - D140_13 * (hpp - hv) / (h13 - hv);
                                                                                if (rtn > Dp6max) rtn = Dp6max;
                                                                                Dp6 = D6(rtn, Qt6, N6);
                                                                                in6 = ogrTG6var(nvar);
                                                                                if (in6 == 1) break;
                                                                                Qpb = Qt - Qpodp- Qt1 - Qt3 - Qt4 - Qt6;
                                                                                rtn = Qpm13 + Qpb / kpb + D13 - rty - D6p(Dp6) - D30_13 * (h30 - hv) / (h13 - hv) - D140_13 * (hpp - hv) / (h13 - hv);
                                                                                Dp2 = DP2(rtn);
                                                                                if (Dp2 > Dp2max) Dp2 = Dp2max;
                                                                                if (Dp2 < Dp2min)
                                                                                {
                                                                                    Dp2 = Dp2min;
                                                                                    if (dp6 != 1)
                                                                                    {
                                                                                        rtn = Qpm13 + Qpb / kpb + D13 - rty - D2p(Dp2) - D30_13 * (h30 - hv) / (h13 - hv) - D140_13 * (hpp - hv) / (h13 - hv);
                                                                                        if (rtn > Dp6max) rtn = Dp6max;
                                                                                        Dp6 = D6(rtn, Qt6, N6);
                                                                                        in6 = ogrTG6var(nvar);
                                                                                        if (in6 == 1) break;
                                                                                    }
                                                                                }
                                                                            }
                                                                            N2 = NN2(Dp2);
                                                                        }
                                                                        in6 = ogrTG6var(nvar);
                                                                        if (in6 == 1) break;
                                                                        if (checkbox30.IsChecked == false) D30_13 = (Qpm13 + Qpb / kpb + D13 - rty - D2p(Dp2) - D6p(Dp6) - D140_13 * (hpp - hv) / (h13 - hv)) / (h30 - hv) * (h13 - hv);
                                                                        if (D30_13 > 120) D30_13 = 120; if (D30_13 < 0) D30_13 = 0;
                                                                        if (checkbox5.IsChecked == false) D140_13 = (Qpm13 + Qpb / kpb + D13 - rty - D2p(Dp2) - D6p(Dp6) - D30_13 * (h30 - hv) / (h13 - hv)) / (hpp - hv) * (h13 - hv);
                                                                        if (D140_13 < 0) D140_13 = 0;
                                                                        rtn = Qpm13 + Qpb / kpb + D13 - rty - D2p(Dp2) - D6p(Dp6) - D30_13 * (h30 - hv) / (h13 - hv) - D140_13 * (hpp - hv) / (h13 - hv);
                                                                        if (Math.Abs(rtn)>2)
                                                                        { 
                                                                            in6 = 1; break; 
                                                                        }
                                                                        if (dp5 == 1) Dp5 = 105;
                                                                        if (dp5 == 2)
                                                                        {
                                                                            Dp5 = Dp30 + D30_13 + Dptn1 + Dptn2 + D3013;
                                                                            if (Dp5 > Dp5max) Dp5 = Dp5max;
                                                                            if (Dp5 < 105)
                                                                            {
                                                                                Dp5 = 105;
                                                                                D30_13 = D30_13 + 105 - Dp5;
                                                                            }
                                                                        }
                                                                        h30n = h5(Dp5);

                                                                        if (checkbox6.IsChecked == false) D140_30 = Dp30 + D30_13 + Dptn1 + Dptn2 + D3013 - Dp5;

                                                                        N2 = NN2(Dp2); N5 = NN5(Dp5);

                                                                        Qpb = Qt - Qpodp- Qt1 - Qt3 - Qt4 - Qt6;

                                                                        if (np1 == 3)
                                                                        {
                                                                            N1 = Ns - N2 - N3 - N4 - N5 - N6;                                                                            
                                                                            Nv1 = ogrTG1n(P1);                                                                            
                                                                            if (tp1 != 4 && N1 < Nv1 && Qt1 > ogr_TG1Qtl(P1, N1) & ch4 != 0) { Qt1 = ogr_TG1Qtl(P1, N1); P1 = Pn(Qt1 * 1000 / (Gsetv1 + 1) + tos1 + dtsp1); if (P1 > 2.5) P1 = 2.5; if (P1 < 0.3) P1 = 0.3; }
                                                                            if (tp1 == 4 && N1 < Nv1 && Qt1 > ogr_TG1Qtl(P1, N1) & ch4 != 0) { in6 = 1; break; }
                                                                            D1t = D1(P1, Qt1, N1);
                                                                            D1t = D1(P1, Qt1, N1 - PoprTG1t(t1, D1t));
                                                                            if (D1t < 0) D1t = 0;
                                                                        }
                                                                        if (np3 == 3)
                                                                        {
                                                                            N3 = Ns - N2 - N1 - N4 - N5 - N6;
                                                                            Nv3 = ogrTG3n(P3, Qt3);
                                                                            if (tp3 != 4 && N3 < Nv3 && Qt3 > ogr_TG3Qtl(P3, N3) & ch4 != 0) { Qt3 = ogr_TG3Qtl(P3, N3); P3 = Pn(Qt3 * 1000 / (Gsetv3 + 1) + tos3 + dtsp3); if (P3 > 2.5) P3 = 2.5; if (P3 < 0.3) P3 = 0.3; }
                                                                            if (tp3 == 4 && N3 < Nv3 && Qt3 > ogr_TG3Qtl(P3, N3) & ch4 != 0) { in6 = 1; break; }
                                                                        }
                                                                        if (np4 == 3)
                                                                        {
                                                                            N4 = Ns - N2 - N1 - N3 - N5 - N6;
                                                                            Nv4 = ogrTG4n(P4, Qt4);
                                                                            if (tp4 != 4 && N4 < Nv4 & Qt4 > ogr_TG4Qtl(P4, N4) & ch4 != 0) { Qt4 = ogr_TG4Qtl(P4, N4); P4 = Pn(Qt4 * 1000 / (Gsetv4 + 1) + tos4 + dtsp4); if (P4 > 2.5) P4 = 2.5; if (P4 < 0.3) P4 = 0.3; }
                                                                            if (tp4 == 4 && N4 < Nv4 & Qt4 > ogr_TG4Qtl(P4, N4) & ch4 != 0) { in6 = 1; break; }
                                                                        }
                                                                        if (np6 == 3)
                                                                        {
                                                                            N6 = Ns - N1 - N3 - N4 - N2 - N5;
                                                                            Nv6 = ogrTG6n(Dp6);
                                                                            if (tp6 != 4 && N6 < Nv6 & Qt6 > ogr_TG6Qtln(Dp6, N6) & ch4 != 0) { Qt6 = ogr_TG6Qtln(Dp6, N6); P6 = Pn(Qt6 * 1000 / (Gsetv6 + 1) + tos6 + dtsp6); if (P6 > 2.5) P6 = 2.5; if (P6 < 0.3) P6 = 0.3; }
                                                                            if (tp6 == 4 && N6 < Nv6 & Qt6 > ogr_TG6Qtln(Dp6, N6) & ch4 != 0) { in6 = 1; break; }
                                                                        }

                                                                        if ((N1 > N1max || N3 > N3max || N4 > N4max || N6 > N6max) & ch4 != 0) { in6 = 1; break; }

                                                                        if (indtg1 == 1)
                                                                        {
                                                                            D3013 = (D1t * (h1 - h13) - (Dhptn(Dptn1) + Dhptn(Dptn2)) + (Dptn1 + Dptn2) * h13) / (h30n - h13);
                                                                            if (D3013 + Dptn1 + Dptn2 > D1t)
                                                                            {
                                                                                D3013 = ((Dptn1 + Dptn2) * h1 - Dhptn(Dptn1) - Dhptn(Dptn2)) / (h30n - h13);
                                                                            }
                                                                        }
                                                                        rty = (Dptn1 + Dptn2 + D3013 - D1t) * (h1 - hv) / (h13 - hv);
                                                                        Qpb = Qt - Qpodp- Qt1 - Qt3 - Qt4 - Qt6;

                                                                        if (indtg1 == 1 & indcn1 == 0)
                                                                        {
                                                                            if (Dk1(P1, Qt1, N1) > 50) indcn1 = 2;
                                                                            else indcn1 = 1;
                                                                        }
                                                                        if (indtg3 == 1 & indcn3 == 0)
                                                                        {
                                                                            if (Dk3(P3, Qt3, N3) > 50) indcn3 = 2;
                                                                            else indcn3 = 1;
                                                                        }
                                                                        if (indtg4 == 1 & indcn4 == 0)
                                                                        {
                                                                            if (Dk4(P4, Qt4, N4) > 100) indcn4 = 2;
                                                                            else indcn4 = 1;
                                                                        }

                                                                        if (indtg4 == 1) { if (indcn4 == 1) Gcw4 = 8000; if (indcn4 == 2) Gcw4 = 16000; }
                                                                        if (indtg3 == 1) { if (indcn3 == 1) Gcw3 = 3500; if (indcn3 == 2) Gcw3 = 7000; }
                                                                        if (indtg1 == 1) { if (indcn1 == 1) Gcw1 = 3500; if (indcn1 == 2) Gcw1 = 7000; }

                                                                        D13 = Daux();
                                                                        ch4++;

                                                                    } while (Math.Abs(D13 - D131) > 3);
                                                                    if (in6 == 1) continue;
                                                                    if (checkBox10.IsChecked == true & Math.Abs(Ns - N1 - N3 - N4 - N2 - N5 - N6) > 1) continue;
                                                                    if (Dp6 < -2 || Dp2 < -2 || D30_13 < -2 || D140_13 < -2 || D140_30 < -2 || Qpb < 0) continue;

                                                                    Gek = (Q2(Dp2) + Q3(P3, Qt3, N3) + Q4(P4, Qt4, N4) + Q5(Dp5) + Q6(Dp6, Qt6, N6)) / 0.595 + D140_13 + D140_30 + D140_45;
                                                                    nk = nek(Gek, im, out G1p, out G2p, out G3p, out G4p, out G5p, out G6p, out G7p);

                                                                    if (nk == 0 || nk == 100) continue;

                                                                    if (N3 > 0 & N3 < 40 & Qt3 > 0)
                                                                    {
                                                                        N3 = N3;
                                                                    }

                                                                    B = Fin(out S, out Nsn, out Kotrt, out Kotrk, out nnet, out ntp, out q);
                                                                    MM = Cn * (N1 + NN2(Dp2) + N3 + N4 + NN5(Dp5) + N6) + (Cgv * Qt + C13 * Qpm13 + C30 * Dp30 + C45 * D140_45) - Ct * (((Q2(Dp2) + Q3(P3, Qt3, N3) + Q4(P4, Qt4, N4) + Q5(Dp5) + Q6(Dp6, Qt6, N6)) + (D140_13 + D140_30 + D140_45) * 0.595) /
                                                                    (7 * 0.98 * nk / 100) + Qptvm(Qpvk) / (7 * 0.98)) - Cnsn * Nsn;

                                                                    if (bteh == 0) MMM = (q * Kotrt) / (7 * Kotrk * nnet * ntp / 10000);
                                                                    else { bte = bteh; MMM = (B - bte * Qs / 1000) / (S - Nsn) * 1000; }

                                                                    if (((radioButton9.IsChecked == true & MM > MMM2) || (radioButton7.IsChecked == true & 1 / MMM > MMM2)) & nk != 0 & nk != 100)
                                                                    {
                                                                        chh = ch3;
                                                                        if (radioButton9.IsChecked == true) MMM2 = MM; if (radioButton7.IsChecked == true) MMM2 = 1 / MMM;
                                                                        Y[0] = Dp6; Y[1] = Qt6; Y[2] = N6; Y[3] = Qt4; Y[4] = N4; Y[5] = Qt3; Y[6] = N3; Y[7] = N1; Y[8] = Qt1; Y[9] = Dp2; Y[10] = Dp5; Y[11] = D140_30; Y[12] = D30_13; Y[13] = D140_13; Y[14] = Dptn1;
                                                                        Y[15] = Dptn2; Y[16] = tp1; Y[17] = np1; Y[18] = tp3; Y[19] = np3; Y[20] = tp4; Y[21] = np4; Y[22] = tp6; Y[23] = np6; Y[24] = dp2; Y[25] = dp6; Y[26] = indptn1; Y[27] = indptn1; Y[28] = Qpb; Y[29] = ch3; Y[30] = nk;
                                                                        Y[31] = P1; Y[32] = P3; Y[33] = P4; Y[34] = P6; Y[35] = indcn1; Y[36] = indcn3; Y[37] = indcn4; Y[38] = indP3; Y[39] = indP4; Y[40] = indP6; Y[41] = D3013; Y[42] = Gsetv1; Y[43] = Gsetv3; Y[44] = Gsetv4; Y[45] = Gsetv6;
                                                                        Y[46] = MM; Y[47] = 1 / MMM; Y[48] = MMM2; Y[49] = D1t; Y[50] = Ddvd; 
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            if (MMM2 == -1000000000000 & combobox10.Text == "неактивен")
            {
                MessageBox.Show("Режим неосуществим", "ОШИБКА", MessageBoxButton.OK, MessageBoxImage.Stop); tabPage1.Focus();
                for (int i = 0; i < 49; i++)
                {
                    Y[i] = 0;
                }
            }
            return Y;
        }

        public double Q6(double D, double Qt, double N)
        {
            double Q = 0;
            if (indtg6 == 1)
            {
                double[] X6 = { 1, D, Qt, N, D * Qt, Qt * N, D * N, D * Qt * N };
                if (indP6 == 1)
                {
                    for (int i = 0; i < 8; i++)
                    {
                        Q = Q + X6[i] * DSTR6[i];
                    }
                }
                else
                {
                    for (int i = 0; i < 8; i++)
                    {
                        Q = Q + X6[i] * DSTR7[i];
                    }
                }
            }
            return Q;
        }
        public double dQ6dDp6(double D, double Qt, double N)
        {
            double Q = 0;
            if (indtg6 == 1)
            {
                double[] X6 = { 0, 1, 0, 0, 2*D, 0, 0, Qt, 0, N, Qt * N, 2*D*Math.Pow(Qt, 2), 0, 2*D*Math.Pow(N, 2), 
            2*D*Math.Pow(Qt, 2)*Math.Pow(N, 2),2*D*Qt, Math.Pow(Qt, 2), 0, 0, 2*D*N, Math.Pow(N, 2), 2*D*Qt*N,
            2*D*Math.Pow(Qt, 2)*N, Math.Pow(Qt, 2)*N, Math.Pow(Qt, 2)*Math.Pow(N, 2), Qt*Math.Pow(N, 2), 2*D*Qt*Math.Pow(N, 2)};
                if (indP6 == 1)
                {
                    for (int i = 0; i < 27; i++)
                    {
                        Q = Q + X6[i] * DSTR6[i];
                    }
                }
                else
                {
                    for (int i = 0; i < 27; i++)
                    {
                        Q = Q + X6[i] * DSTR7[i];
                    }
                }
            }
            return Q;
        }
        public double dQ6dQt6(double D, double Qt, double N)
        {
            double Q = 0;
            if (indtg6 == 1)
            {
                double[] X6 = { 0, 0, 1, 0, 0, 2*Qt, 0, D, N, 0, D * N, Math.Pow(D, 2)*2*Qt, 2*Qt*Math.Pow(N, 2),
            0, Math.Pow(D, 2)*2*Qt*Math.Pow(N, 2), Math.Pow(D, 2), D*2*Qt, 2*Qt*N, Math.Pow(N, 2),
            0, 0, Math.Pow(D, 2)*N, Math.Pow(D, 2)*2*Qt*N, D*2*Qt*N, D*2*Qt*Math.Pow(N, 2),
            D*Math.Pow(N, 2), Math.Pow(D, 2)*Math.Pow(N, 2)};
                if (indP6 == 1)
                {
                    for (int i = 0; i < 27; i++)
                    {
                        Q = Q + X6[i] * DSTR6[i];
                    }
                }
                else
                {
                    for (int i = 0; i < 27; i++)
                    {
                        Q = Q + X6[i] * DSTR7[i];
                    }
                }
            }
            return Q;
        }
        public double dQ6dN6(double D, double Qt, double N)
        {
            double Q = 0;
            if (indtg6 == 1)
            {
                double[] X6 = { 0, 0, 0, 1, 0, 0, 2*N, 0, Qt, D, D * Qt, 0, Math.Pow(Qt, 2)*2*N, Math.Pow(D, 2)*2*N, 
            Math.Pow(D, 2)*Math.Pow(Qt, 2)*2*N, 0, 0, Math.Pow(Qt, 2), Qt*2*N, Math.Pow(D, 2), D*2*N,
            Math.Pow(D, 2)*Qt, Math.Pow(D, 2)*Math.Pow(Qt, 2), D*Math.Pow(Qt, 2), D*Math.Pow(Qt, 2)*2*N,
            D*Qt*2*N, Math.Pow(D, 2)*Qt*2*N};
                if (indP6 == 1)
                {
                    for (int i = 0; i < 27; i++)
                    {
                        Q = Q + X6[i] * DSTR6[i];
                    }
                }
                else
                {
                    for (int i = 0; i < 27; i++)
                    {
                        Q = Q + X6[i] * DSTR7[i];
                    }
                }
            }
            return Q;
        }
        public double dQ6dN6dN6(double D, double Qt)
        {
            double Q = 0;
            if (indtg6 == 1)
            {
                double[] X6 =  { 0, 0, 0, 0, 0, 0, 2, 0, 0, 0, 0, 0, Math.Pow(Qt, 2)*2, Math.Pow(D, 2)*2, Math.Pow(D, 2)*Math.Pow(Qt, 2)*2, 0, 0, 0, 
            Qt*2, 0, D*2, 0, 0, 0, D*Math.Pow(Qt, 2)*2, D*Qt*2, Math.Pow(D, 2)*Qt*2};
                if (indP6 == 1)
                {
                    for (int i = 0; i < 27; i++)
                    {
                        Q = Q + X6[i] * DSTR6[i];
                    }
                }
                else
                {
                    for (int i = 0; i < 27; i++)
                    {
                        Q = Q + X6[i] * DSTR7[i];
                    }
                }
            }
            return Q;
        }
        public double dQ6dQt6dN6(double D, double Qt, double N)
        {
            double Q = 0;
            if (indtg6 == 1)
            {
                double[] X6 = { 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, D, 0, 2*Qt*2*N, 0, Math.Pow(D, 2)*2*Qt*2*N, 0, 0, 2*Qt, 2*N,
            0, 0, Math.Pow(D, 2), Math.Pow(D, 2)*2*Qt, D*2*Qt, D*2*Qt*2*N, D*2*N, Math.Pow(D, 2)*2*N};
                if (indP6 == 1)
                {
                    for (int i = 0; i < 27; i++)
                    {
                        Q = Q + X6[i] * DSTR6[i];
                    }
                }
                else
                {
                    for (int i = 0; i < 27; i++)
                    {
                        Q = Q + X6[i] * DSTR7[i];
                    }
                }
            }
            return Q;
        }
        public double dQ6dQt6dQt6(double D, double N)
        {
            double Q = 0;
            if (indtg6 == 1)
            {
                double[] X6 = { 0, 0, 0, 0, 0, 2, 0, 0, 0, 0, 0, Math.Pow(D, 2)*2, 2*Math.Pow(N, 2), 0, Math.Pow(D, 2)*2*Math.Pow(N, 2), 0, D*2, 2*N, 0,
            0, 0, 0, Math.Pow(D, 2)*2*N, D*2*N, D*2*Math.Pow(N, 2), 0, 0};
                if (indP6 == 1)
                {
                    for (int i = 0; i < 27; i++)
                    {
                        Q = Q + X6[i] * DSTR6[i];
                    }
                }
                else
                {
                    for (int i = 0; i < 27; i++)
                    {
                        Q = Q + X6[i] * DSTR7[i];
                    }
                }
            }
            return Q;
        }
        public double dQ6dDp6dN6(double D, double Qt, double N)
        {
            double Q = 0;
            if (indtg6 == 1)
            {
                double[] X6 = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, Qt, 0, 0, 2*D*2*N, 2*D*Math.Pow(Qt, 2)*2*N, 0, 0, 0, 0, 2*D, 2*N, 2*D*Qt,
            2*D*Math.Pow(Qt, 2), Math.Pow(Qt, 2), Math.Pow(Qt, 2)*2*N, Qt*2*N, 2*D*Qt*2*N};
                if (indP6 == 1)
                {
                    for (int i = 0; i < 27; i++)
                    {
                        Q = Q + X6[i] * DSTR6[i];
                    }
                }
                else
                {
                    for (int i = 0; i < 27; i++)
                    {
                        Q = Q + X6[i] * DSTR7[i];
                    }
                }
            }
            return Q;
        }
        public double dQ6dDp6dQt6(double D, double Qt, double N)
        {
            double Q = 0;
            if (indtg6 == 1)
            {
                double[] X6 = { 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, N, 2*D*2*Qt, 0, 0, 2*D*2*Qt*Math.Pow(N, 2), 2*D, 2*Qt, 0, 0, 0, 0, 2*D*N,
            2*D*2*Qt*N, 2*Qt*N, 2*Qt*Math.Pow(N, 2), Math.Pow(N, 2), 2*D*Math.Pow(N, 2)};
                if (indP6 == 1)
                {
                    for (int i = 0; i < 27; i++)
                    {
                        Q = Q + X6[i] * DSTR6[i];
                    }
                }
                else
                {
                    for (int i = 0; i < 27; i++)
                    {
                        Q = Q + X6[i] * DSTR7[i];
                    }
                }
            }
            return Q;
        }
        public double dQ6dDp6dDp6(double Qt, double N)
        {
            double Q = 0;
            if (indtg6 == 1)
            {
                double[] X6 =  { 0, 0, 0, 0, 2, 0, 0, 0, 0, 0, 0, 2*Math.Pow(Qt, 2), 0, 2*Math.Pow(N, 2), 2*Math.Pow(Qt, 2)*Math.Pow(N, 2),
            2*Qt, 0, 0, 0, 2*N, 0, 2*Qt*N, 2*Math.Pow(Qt, 2)*N, 0, 0, 0, 2*Qt*Math.Pow(N, 2)};
                if (indP6 == 1)
                {
                    for (int i = 0; i < 27; i++)
                    {
                        Q = Q + X6[i] * DSTR6[i];
                    }
                }
                else
                {
                    for (int i = 0; i < 27; i++)
                    {
                        Q = Q + X6[i] * DSTR7[i];
                    }
                }
            }
            return Q;
        }

        public double Qp6(double D, double Qt, double N)
        {
            double Q = 0;
            if (D != 0)
            {
                double[] X6 = { 1, D, Qt, N, Math.Pow(D, 2), Math.Pow(Qt, 2), Math.Pow(N, 2), D * Qt, Qt * N, D * N, D * Qt * N, Math.Pow(D, 2) * Math.Pow(Qt, 2), Math.Pow(Qt, 2) * Math.Pow(N, 2), Math.Pow(D, 2) * Math.Pow(N, 2), Math.Pow(D, 2) * Math.Pow(Qt, 2) * Math.Pow(N, 2), Math.Pow(D, 2) * Qt, D * Math.Pow(Qt, 2), Math.Pow(Qt, 2) * N, Qt * Math.Pow(N, 2), Math.Pow(D, 2) * N, D * Math.Pow(N, 2), Math.Pow(D, 2) * Qt * N, Math.Pow(D, 2) * Math.Pow(Qt, 2) * N, D * Math.Pow(Qt, 2) * N, D * Math.Pow(Qt, 2) * Math.Pow(N, 2), D * Qt * Math.Pow(N, 2), Math.Pow(D, 2) * Qt * Math.Pow(N, 2) };
                for (int i = 0; i < 27; i++)
                {
                    Q = Q + X6[i] * DSTR54[i];
                }
            }
            else Q = 0;
            return Q;
        }
        public double DQ6(double D, double Qt, double N)
        {
            double Q = 0;
            if (D != 0)
            {
                double[] X6 = { 1, D, Qt, N, Math.Pow(D, 2), Math.Pow(Qt, 2), Math.Pow(N, 2), D * Qt, Qt * N, D * N, D * Qt * N, Math.Pow(D, 2) * Math.Pow(Qt, 2), Math.Pow(Qt, 2) * Math.Pow(N, 2), Math.Pow(D, 2) * Math.Pow(N, 2), Math.Pow(D, 2) * Math.Pow(Qt, 2) * Math.Pow(N, 2), Math.Pow(D, 2) * Qt, D * Math.Pow(Qt, 2), Math.Pow(Qt, 2) * N, Qt * Math.Pow(N, 2), Math.Pow(D, 2) * N, D * Math.Pow(N, 2), Math.Pow(D, 2) * Qt * N, Math.Pow(D, 2) * Math.Pow(Qt, 2) * N, D * Math.Pow(Qt, 2) * N, D * Math.Pow(Qt, 2) * Math.Pow(N, 2), D * Qt * Math.Pow(N, 2), Math.Pow(D, 2) * Qt * Math.Pow(N, 2) };
                for (int i = 0; i < 27; i++)
                {
                    Q = Q + X6[i] * DSTR51[i];
                }
            }
            else Q = 0;
            return Q;
        }
        public double dDQ6dDp6(double D, double Qt, double N)
        {
            double Q = 0;
            double[] X6 = { 0, 1, 0, 0, 2*D, 0, 0, Qt, 0, N, Qt * N, 2*D*Math.Pow(Qt, 2), 0, 2*D*Math.Pow(N, 2), 2*D*Math.Pow(Qt, 2)*Math.Pow(N, 2), 2*D*Qt, Math.Pow(Qt, 2), 
            0, 0, 2*D*N, Math.Pow(N, 2), 2*D*Qt*N, 2*D*Math.Pow(Qt, 2)*N, Math.Pow(Qt, 2)*N, Math.Pow(Qt, 2)*Math.Pow(N, 2), Qt*Math.Pow(N, 2), 2*D*Qt*Math.Pow(N, 2)};
            for (int i = 0; i < 27; i++)
            {
                Q = Q + X6[i] * DSTR51[i];
            }
            return Q;
        }
        public double dDQ6dQt6(double D, double Qt, double N)
        {
            double Q = 0;
            double[] X6 = { 0, 0, 1, 0, 0, 2*Qt, 0, D, N, 0, D * N, Math.Pow(D, 2)*2*Qt, 2*Qt*Math.Pow(N, 2), 0, Math.Pow(D, 2)*2*Qt*Math.Pow(N, 2),Math.Pow(D, 2), 
                D*2*Qt, 2*Qt*N, Math.Pow(N, 2), 0, 0, Math.Pow(D, 2)*N, Math.Pow(D, 2)*2*Qt*N, D*2*Qt*N, D*2*Qt*Math.Pow(N, 2), D*Math.Pow(N, 2),Math.Pow(D, 2)*Math.Pow(N, 2)};
            for (int i = 0; i < 27; i++)
            {
                Q = Q + X6[i] * DSTR51[i];
            }
            return Q;
        }
        public double dDQ6dN6(double D, double Qt, double N)
        {
            double Q = 0;
            double[] X6 = { 0, 0, 0, 1, 0, 0, 2*N, 0, Qt, D, D * Qt, 0, Math.Pow(Qt, 2)*2*N, Math.Pow(D, 2)*2*N, Math.Pow(D, 2)*Math.Pow(Qt, 2)*2*N, 0, 0, Math.Pow(Qt, 2),
                Qt*2*N, Math.Pow(D, 2), D*2*N, Math.Pow(D, 2)*Qt, Math.Pow(D, 2)*Math.Pow(Qt, 2), D*Math.Pow(Qt, 2), D*Math.Pow(Qt, 2)*2*N, D*Qt*2*N, Math.Pow(D, 2)*Qt*2*N};
            for (int i = 0; i < 27; i++)
            {
                Q = Q + X6[i] * DSTR51[i];
            }
            return Q;
        }
        public double dDQ6dN6dN6(double D, double Qt)
        {
            double Q = 0;
            double[] X6 = { 0, 0, 0, 0, 0, 0, 2, 0, 0, 0, 0, 0, Math.Pow(Qt, 2)*2, Math.Pow(D, 2)*2, Math.Pow(D, 2)*Math.Pow(Qt, 2)*2, 0, 0, 0,
                Qt*2, 0, D*2, 0, 0, 0, D*Math.Pow(Qt, 2)*2, D*Qt*2, Math.Pow(D, 2)*Qt*2};
            for (int i = 0; i < 27; i++)
            {
                Q = Q + X6[i] * DSTR51[i];
            }
            return Q;
        }
        public double dDQ6dQt6dN6(double D, double Qt, double N)
        {
            double Q = 0;
            double[] X6 = { 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, D, 0, 2*Qt*2*N, 0, Math.Pow(D, 2)*2*Qt*2*N,0, 0, 2*Qt, 
            2*N, 0, 0, Math.Pow(D, 2), Math.Pow(D, 2)*2*Qt, D*2*Qt, D*2*Qt*2*N, D*2*N, Math.Pow(D, 2)*2*N};
            for (int i = 0; i < 27; i++)
            {
                Q = Q + X6[i] * DSTR51[i];
            }
            return Q;
        }
        public double dDQ6dQt6dQt6(double D, double N)
        {
            double Q = 0;
            double[] X6 = { 0, 0, 0, 0, 0, 2, 0, 0, 0, 0, 0, Math.Pow(D, 2)*2, 2*Math.Pow(N, 2), 0, Math.Pow(D, 2)*2*Math.Pow(N, 2), 0, 
            D*2, 2*N, 0, 0, 0, 0, Math.Pow(D, 2)*2*N, D*2*N, D*2*Math.Pow(N, 2), 0, 0};
            for (int i = 0; i < 27; i++)
            {
                Q = Q + X6[i] * DSTR51[i];
            }
            return Q;
        }
        public double dDQ6dDp6dN6(double D, double Qt, double N)
        {
            double Q = 0;
            double[] X6 = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, Qt, 0, 0, 2*D*2*N, 2*D*Math.Pow(Qt, 2)*2*N, 0, 0, 0, 
            0, 2*D, 2*N, 2*D*Qt, 2*D*Math.Pow(Qt, 2), Math.Pow(Qt, 2), Math.Pow(Qt, 2)*2*N, Qt*2*N, 2*D*Qt*2*N};
            for (int i = 0; i < 27; i++)
            {
                Q = Q + X6[i] * DSTR51[i];
            }
            return Q;
        }
        public double dDQ6dDp6dQt6(double D, double Qt, double N)
        {
            double Q = 0;
            double[] X6 = { 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, N, 2*D*2*Qt, 0, 0, 2*D*2*Qt*Math.Pow(N, 2), 2*D, 
            2*Qt, 0, 0, 0, 0, 2*D*N, 2*D*2*Qt*N, 2*Qt*N, 2*Qt*Math.Pow(N, 2), Math.Pow(N, 2), 2*D*Math.Pow(N, 2)};
            for (int i = 0; i < 27; i++)
            {
                Q = Q + X6[i] * DSTR51[i];
            }
            return Q;
        }
        public double dDQ6dDp6dDp6(double Qt, double N)
        {
            double Q = 0;
            double[] X6 = { 0, 0, 0, 0, 2, 0, 0, 0, 0, 0, 0, 2*Math.Pow(Qt, 2), 0, 2*Math.Pow(N, 2), 2*Math.Pow(Qt, 2)*Math.Pow(N, 2), 2*Qt, Math.Pow(Qt, 2), 
                0, 0, 2*N, 0, 2*Qt*N, 2*Math.Pow(Qt, 2)*N, 0, 0, 0, 2*Qt*Math.Pow(N, 2)};
            for (int i = 0; i < 27; i++)
            {
                Q = Q + X6[i] * DSTR51[i];
            }
            return Q;
        }

        public double Q4(double P, double Qt, double N)
        {
            double Q = 0;
            if (indtg4 == 1)
            {
                double[] X4 = { 1, P, Qt, N, P * Qt, Qt * N, P * N, P * Qt * N };
                if (indP4 == 1)
                {
                    for (int i = 0; i < 8; i++)
                    {
                        Q = Q + X4[i] * DSTR4[i];
                    }
                }
                else
                {
                    for (int i = 0; i < 8; i++)
                    {
                        Q = Q + X4[i] * DSTR5[i];
                    }
                }
            }
            return Q;
        }
        public double dQ4dN4(double P, double Qt, double N)
        {
            double Q = 0;
            double[] X4 = { 0, 0, 0, 1, 0, 0, 2*N, 0, Qt, P, P * Qt, 0, Math.Pow(Qt, 2)*2*N, Math.Pow(P, 2)*2*N, Math.Pow(P, 2)*Math.Pow(Qt, 2)*2*N, 0, 0, 
            Math.Pow(Qt, 2), Qt*2*N, Math.Pow(P, 2), P*2*N, Math.Pow(P, 2)*Qt,Math.Pow(P, 2)*Math.Pow(Qt, 2),P*Math.Pow(Qt, 2),
            P*Math.Pow(Qt, 2)*2*N,P*Qt*2*N,Math.Pow(P, 2)*Qt*2*N};
            if (indP4 == 1)
            {
                for (int i = 0; i < 27; i++)
                {
                    Q = Q + X4[i] * DSTR4[i];
                }
            }
            else
            {
                for (int i = 0; i < 27; i++)
                {
                    Q = Q + X4[i] * DSTR5[i];
                }
            }
            return Q;
        }
        public double dQ4dN4dN4(double P, double Qt)
        {
            double Q = 0;
            double[] X4 = { 0, 0, 0, 0, 0, 0, 2, 0, 0, 0, 0, 0, Math.Pow(Qt, 2)*2, Math.Pow(P, 2)*2, Math.Pow(P, 2)*Math.Pow(Qt, 2)*2, 0, 0, 
            0, Qt*2, 0, P*2, 0, 0, 0, P*Math.Pow(Qt, 2)*2, P*Qt*2,Math.Pow(P, 2)*Qt*2};
            if (indP4 == 1)
            {
                for (int i = 0; i < 27; i++)
                {
                    Q = Q + X4[i] * DSTR4[i];
                }
            }
            else
            {
                for (int i = 0; i < 27; i++)
                {
                    Q = Q + X4[i] * DSTR5[i];
                }
            }
            return Q;
        }
        public double dQ4dN4dQt4(double P, double Qt, double N)
        {
            double Q = 0;
            double[] X4 = { 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, P, 0, 2*Qt*2*N, 0, Math.Pow(P, 2)*2*Qt*2*N, 0, 0, 2*Qt, 2*N, 0, 0,
            Math.Pow(P, 2), Math.Pow(P, 2)*2*Qt, P*2*Qt, P*2*Qt*2*N, P*2*N, Math.Pow(P, 2)*2*N};
            if (indP4 == 1)
            {
                for (int i = 0; i < 27; i++)
                {
                    Q = Q + X4[i] * DSTR4[i];
                }
            }
            else
            {
                for (int i = 0; i < 27; i++)
                {
                    Q = Q + X4[i] * DSTR5[i];
                }
            }
            return Q;
        }
        public double dQ4dQt4(double P, double Qt, double N)
        {
            double Q = 0;
            double[] X4 = { 0, 0, 1, 0, 0, 2*Qt, 0, P , N, 0, P * N, Math.Pow(P, 2)*2*Qt, 2*Qt*Math.Pow(N, 2), 0, 
            Math.Pow(P, 2)*2*Qt*Math.Pow(N, 2),Math.Pow(P, 2), P*2*Qt, 2*Qt*N, Math.Pow(N, 2), 0, 0, Math.Pow(P, 2)*N,
            Math.Pow(P, 2)*2*Qt*N, P*2*Qt*N, P*2*Qt*Math.Pow(N, 2),P*Math.Pow(N, 2), Math.Pow(P, 2)*Math.Pow(N, 2)};
            if (indP4 == 1)
            {
                for (int i = 0; i < 27; i++)
                {
                    Q = Q + X4[i] * DSTR4[i];
                }
            }
            else
            {
                for (int i = 0; i < 27; i++)
                {
                    Q = Q + X4[i] * DSTR5[i];
                }
            }
            return Q;
        }
        public double dQ4dQt4dQt4(double P, double N)
        {
            double Q = 0;
            double[] X4 = { 0, 0, 0, 0, 0, 2, 0, 0, 0, 0, 0, Math.Pow(P, 2)*2, 2*Math.Pow(N, 2), 0, Math.Pow(P, 2)*2*Math.Pow(N, 2), 0, P*2, 2*N, 0, 0, 0, 0,
            Math.Pow(P, 2)*2*N, P*2*N, P*2*Math.Pow(N, 2), 0, 0};
            if (indP4 == 1)
            {
                for (int i = 0; i < 27; i++)
                {
                    Q = Q + X4[i] * DSTR4[i];
                }
            }
            else
            {
                for (int i = 0; i < 27; i++)
                {
                    Q = Q + X4[i] * DSTR5[i];
                }
            }
            return Q;
        }

        public double Q3(double P, double Qt, double N)
        {
            double Q = 0;
            if (indtg3 == 1)
            {
                double[] X3 = {1, P, Qt, N, P * Qt, Qt * N, P * N, P * Qt * N} ;
                if (indP3 == 1)
                {
                    for (int i = 0; i < 8; i++)
                    {
                        Q = Q + X3[i] * DSTR2[i];
                    }
                }
                else
                {
                    for (int i = 0; i < 8; i++)
                    {
                        Q = Q + X3[i] * DSTR3[i];
                    }
                }
            }
            return Q;
        }
        public double dQ3dN3(double P, double Qt, double N)
        {
            double Q = 0;
            double[] X3 = { 0, 0, 0, 1, 0, 0, 2*N, 0, Qt, P, P * Qt, 0, Math.Pow(Qt, 2)*2*N, Math.Pow(P, 2)*2*N, Math.Pow(P, 2)*Math.Pow(Qt, 2)*2*N, 0, 0, 
            Math.Pow(Qt, 2), Qt*2*N, Math.Pow(P, 2), P*2*N, Math.Pow(P, 2)*Qt,Math.Pow(P, 2)*Math.Pow(Qt, 2),P*Math.Pow(Qt, 2),
            P*Math.Pow(Qt, 2)*2*N,P*Qt*2*N,Math.Pow(P, 2)*Qt*2*N};
            if (indP3 == 1)
            {
                for (int i = 0; i < 27; i++)
                {
                    Q = Q + X3[i] * DSTR2[i];
                }
            }
            else
            {
                for (int i = 0; i < 27; i++)
                {
                    Q = Q + X3[i] * DSTR3[i];
                }
            }
            return Q = 0;
        }
        public double dQ3dNdN3(double P, double Qt)
        {
            double Q = 0;
            double[] X3 = { 0, 0, 0, 0, 0, 0, 2, 0, 0, 0, 0, 0, Math.Pow(Qt, 2)*2, Math.Pow(P, 2)*2, Math.Pow(P, 2)*Math.Pow(Qt, 2)*2, 0, 0, 
            0, Qt*2, 0, P*2, 0, 0, 0, P*Math.Pow(Qt, 2)*2, P*Qt*2,Math.Pow(P, 2)*Qt*2};
            if (indP3 == 1)
            {
                for (int i = 0; i < 27; i++)
                {
                    Q = Q + X3[i] * DSTR2[i];
                }
            }
            else
            {
                for (int i = 0; i < 27; i++)
                {
                    Q = Q + X3[i] * DSTR3[i];
                }
            }
            return Q = 0;
        }
        public double dQ3dQt3dN3(double P, double Qt, double N)
        {
            double Q = 0;
            double[] X3 =  { 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, P, 0, 2*Qt*2*N, 0, Math.Pow(P, 2)*2*Qt*2*N, 0, 0, 2*Qt, 2*N, 0, 0, 
            Math.Pow(P, 2), Math.Pow(P, 2)*2*Qt, P*2*Qt, P*2*Qt*2*N, P*2*N, Math.Pow(P, 2)*2*N};
            if (indP3 == 1)
            {
                for (int i = 0; i < 27; i++)
                {
                    Q = Q + X3[i] * DSTR2[i];
                }
            }
            else
            {
                for (int i = 0; i < 27; i++)
                {
                    Q = Q + X3[i] * DSTR3[i];
                }
            }
            return Q;
        }
        public double dQ3dQt3(double P, double Qt, double N)
        {
            double Q = 0;
            double[] X3 =  { 0, 0, 1, 0, 0, 2*Qt, 0, P , N, 0, P * N, Math.Pow(P, 2)*2*Qt, 2*Qt*Math.Pow(N, 2), 0, 
            Math.Pow(P, 2)*2*Qt*Math.Pow(N, 2),Math.Pow(P, 2), P*2*Qt, 2*Qt*N, Math.Pow(N, 2), 0, 0, Math.Pow(P, 2)*N,
            Math.Pow(P, 2)*2*Qt*N, P*2*Qt*N, P*2*Qt*Math.Pow(N, 2),P*Math.Pow(N, 2), Math.Pow(P, 2)*Math.Pow(N, 2)};
            if (indP3 == 1)
            {
                for (int i = 0; i < 27; i++)
                {
                    Q = Q + X3[i] * DSTR2[i];
                }
            }
            else
            {
                for (int i = 0; i < 27; i++)
                {
                    Q = Q + X3[i] * DSTR3[i];
                }
            }
            return Q;
        }
        public double dQ3dQt3dQt3(double P, double N)
        {
            double Q = 0;
            double[] X3 =  { 0, 0, 0, 0, 0, 2, 0, 0, 0, 0, 0, Math.Pow(P, 2)*2, 2*Math.Pow(N, 2), 0, Math.Pow(P, 2)*2*Math.Pow(N, 2), 0, P*2, 2*N, 0, 0, 0, 0,
            Math.Pow(P, 2)*2*N, P*2*N, P*2*Math.Pow(N, 2), 0, 0};
            if (indP3 == 1)
            {
                for (int i = 0; i < 27; i++)
                {
                    Q = Q + X3[i] * DSTR2[i];
                }
            }
            else
            {
                for (int i = 0; i < 27; i++)
                {
                    Q = Q + X3[i] * DSTR3[i];
                }
            }
            return Q;
        }

        public double Q2(double Q)
        {
            double Q2 = 0;
            if (indtg2 == 1)
            {
                double[] X2 = { 1, Q };
                for (int i = 0; i < 2; i++)
                {
                    Q2 = Q2 + X2[i] * DSTR52[i];
                }
            }
            return Q2;
        }
        public double dQ2dDp2()
        {
            double Q = 0;
            double[] X2 = { 0, 1 };
            for (int i = 0; i < 2; i++)
            {
                Q = Q + X2[i] * DSTR52[i];
            }
            return Q;
        }
        public double NN2(double D)
        {
            double N = 0;
            if (D != 0)
            {
                double[] X2 = { 1, D };
                for (int i = 0; i < 2; i++)
                {
                    N = N + X2[i] * DSTR45[i];
                }
            }
            return N;
        }
        public double dNN2dDp()
        {
            double N = 0;
            double[] X2 = { 0, 1 };
            for (int i = 0; i < 2; i++)
            {
                N = N + X2[i] * DSTR45[i];
            }
            return N;
        }

        public double NN5(double D)
        {
            double N = 0;
            if (D != 0)
            {
                double[] X5 = { 1, D };
                for (int i = 0; i < 2; i++)
                {
                    N = N + X5[i] * DSTR46[i];
                }
                if (N < 0) N = 0;
            }
            return N;
        }

        public double dFdDp2()
        {
            double F = 0;
            F = Cn * dNN2dDp() - Ct * (dQ2dDp2() / (7 * 0.98 * nk / 100)) - Cnsn * (dNsnpndDp2() + dNsnko() * dGpvdDp2());
            return F;
        }

        public double dFdN1()
        {
            double F = Cn - Cnsn * dNsntodN1();
            return F;
        }
        public double dFdN1dN1()
        {
            double F = -Cnsn * dNsntodN1dN1();
            return F;
        }

        public double dFdN3()
        {
            double F = 0;
            F = Cn - Ct * dQ3dN3(P3, Qt3, N3) / (7 * 0.98 * nk / 100) - Cnsn * (dNsntodN3() + dNsnpndN3() + dNsnko() * dGpvdN3());
            return F;
        }
        public double dFdN3dN3()
        {
            double F = -Ct * dQ3dNdN3(P3, Qt3) / (7 * 0.98 * nk / 100) - Cnsn * (dNsntodN3dN3() + dNsnpndN3dN3() + dNsnko() * dGpvdN3dN3());
            return F;
        }
        public double dFdQt3()
        {
            double F = 0;
            F = -Ct * dQ3dQt3(P3, Qt3, N3) / (7 * 0.98 * nk / 100) - Cnsn * (dNsnpndQt3() + dNsnko() * dGpvdQt3());
            return F;
        }
        public double dFdQt3dQt3()
        {
            double F = -Ct * dQ3dQt3dQt3(P3, N3) / (7 * 0.98 * nk / 100) - Cnsn * (dNsnpndQt3dQt3() + dNsnko() * dGpvdQt3dQt3());
            return F;
        }
        public double dFdQt3dN3()
        {
            double F = -Ct * dQ3dQt3dN3(P3, Qt3, N3) / (7 * 0.98 * nk / 100) - Cnsn * (dNsnpndQt3dN3() + dNsnko() * dGpvdQt3dN3());
            return F;
        }

        public double dFdN4()
        {
            double F = 0;
            F = Cn - Ct * dQ4dN4(P4, Qt4, N4) / (7 * 0.98 * nk / 100) - Cnsn * (dNsntodN4() + dNsnpndN4() + dNsnko() * dGpvdN4());
            return F;
        }
        public double dFdN4dN4()
        {
            double F = -Ct * dQ4dN4dN4(P4, Qt4) / (7 * 0.98 * nk / 100) - Cnsn * (dNsntodN4dN4() + dNsnpndN4dN4() + dNsnko() * dGpvdN4dN4());
            return F;
        }
        public double dFdQt4()
        {
            double F = 0;
            F = -Ct * dQ4dQt4(P4, Qt4, N4) / (7 * 0.98 * nk / 100) - Cnsn * (dNsnpndQt4() + dNsnko() * dGpvdQt4());
            return F;
        }
        public double dFdQt4dN4()
        {
            double F = -Ct * dQ4dN4dQt4(P4, Qt4, N4) / (7 * 0.98 * nk / 100) - Cnsn * (dNsnpndQt4dN4() + dNsnko() * dGpvdQt4dN4());
            return F;
        }
        public double dFdQt4dQt4()
        {
            double F = -Ct * dQ4dQt4dQt4(P4, N4) / (7 * 0.98 * nk / 100) - Cnsn * (dNsnpndQt4dQt4() + dNsnko() * dGpvdQt4dQt4());
            return F;
        }

        public double dFdN6()
        {
            double F = 0;
            F = Cn - Ct * dQ6dN6(Dp6, Qt6, N6) / (7 * 0.98 * nk / 100) - Cnsn * (dNsntodN6() + dNsnpndN6() + dNsnko() * dGpvdN6());
            return F;
        }
        public double dFdN6dN6()
        {
            double F = -Ct * dQ6dN6dN6(Dp6, Qt6) / (7 * 0.98 * nk / 100) - Cnsn * (dNsntodN6dN6() + dNsnpndN6dN6() + dNsnko() * dGpvdN6dN6());
            return F;
        }
        public double dFdQt6()
        {
            double F = -Ct * dQ6dQt6(Dp6, Qt6, N6) / (7 * 0.98 * nk / 100) - Cnsn * (dNsnpndQt6() + dNsnko() * dGpvdQt6());
            return F;
        }
        public double dFdQt6dN6()
        {
            double F = -Ct * dQ6dQt6dN6(Dp6, Qt6, N6) / (7 * 0.98 * nk / 100) - Cnsn * (dNsnpndQt6dN6() + dNsnko() * dGpvdQt6dN6());
            return F;
        }
        public double dFdQt6dQt6()
        {
            double F = -Ct * dQ6dQt6dQt6(Dp6, N6) / (7 * 0.98 * nk / 100) - Cnsn * (dNsnpndQt6dQt6() + dNsnko() * dGpvdQt6dQt6());
            return F;
        }
        public double dFdDp6()
        {
            double F = -Ct * dQ6dDp6(Dp6, Qt6, N6) / (7 * 0.98 * nk / 100) - Cnsn * (dNsnpndDp6() + dNsnko() * dGpvdDp6());
            return F;
        }
        public double dFdDp6dN6()
        {
            double F = -Ct * dQ6dDp6dN6(Dp6, Qt6, N6) / (7 * 0.98 * nk / 100) - Cnsn * (dNsnpndDp6dN6() + dNsnko() * dGpvdDp6dN6());
            return F;
        }
        public double dFdDp6dQt6()
        {
            double F = -Ct * dQ6dDp6dQt6(Dp6, Qt6, N6) / (7 * 0.98 * nk / 100) - Cnsn * (dNsnpndDp6dQt6() + dNsnko() * dGpvdDp6dQt6());
            return F;
        }
        public double dFdDp6dDp6()
        {
            double F = -Ct * dQ6dDp6dDp6(Qt6, N6) / (7 * 0.98 * nk / 100) - Cnsn * (dNsnpndDp6dDp6() + dNsnko() * dGpvdDp6dDp6());
            return F;
        }

        public double dNsnpndN6()
        {
            double npn = 0.0085088 * (dQ6dN6(Dp6, Qt6, N6) / 0.595 / 0.995);
            return npn;
        }
        public double dNsnpndDp6()
        {
            double npn = 0.0085088 * (dQ6dDp6(Dp6, Qt6, N6) / 0.595 / 0.995);
            return npn;
        }
        public double dNsnpndQt6()
        {
            double npn;
            npn = 0.0085088 * dQ6dQt6(Dp6, Qt6, N6) / 0.595 / 0.995;
            return npn;
        }
        public double dNsnpndQt4()
        {
            double npn;
            npn = 0.0085088 * dQ4dQt4(P4, Qt4, N4) / 0.595 / 0.995;
            return npn;
        }
        public double dNsnpndN4()
        {
            double npn = 0.0085088 * (dQ4dN4(P4, Qt4, N4) / 0.595 / 0.995);
            return npn;
        }
        public double dNsnpndDp2()
        {
            double npn = 0.0085088 * (dQ2dDp2() / 0.595 / 0.995);
            return npn;
        }
        public double dNsnpndN3()
        {
            double npn = 0.0085088 * (dQ3dN3(P3, Qt3, N3) / 0.595 / 0.995);
            return npn;
        }
        public double dNsnpndQt3()
        {
            double npn = 0.0085088 * dQ3dQt3(P3, Qt3, N3) / 0.595 / 0.995;
            return npn;
        }
        public double dNsnpndDp5()
        {
            double npn;
            npn = 0.0085088 * (dQ5dD5() / 0.595 / 0.995);
            return npn;
        }
        public double dNsnpndDptn()
        {
            double npn = -0.0085088 * 10;
            return npn;
        }

        public double dNsnko()
        {
            double snk = 0.0005 * 0.595 + 0.6528;
            return snk;
        }
        public double dGpvdDp6()
        {
            double G = dQ6dDp6(Dp6, Qt6, N6) / 0.595 / 0.995;
            return G;
        }
        public double dNsntodN6()
        {
            double sn = 0;
            sn = -0.0000211825 * 2 * N6 + 0.0048747558;
            return sn;
        } //N6         
        public double dGpvdN6()
        {
            double G = dQ6dN6(Dp6, Qt6, N6) / 0.595 / 0.995;
            return G;
        }
        public double dGpvdN4()
        {
            double G = dQ4dN4(P4, Qt4, N4) / 0.595 / 0.995;
            return G;
        }
        public double dNsntodN4()
        {
            double sn = -0.0000134196 * 2 * N4 + 0.0030831795;
            return sn;
        } //N4        
        public double dGpvdDp2()
        {
            double G = dQ2dDp2() / 0.595 / 0.995;
            return G;
        }
        public double dGpvdQt4()
        {
            double G = dQ4dQt4(P4, Qt4, N4) / 0.595 / 0.995;
            return G;
        }
        public double dGpvdQt6()
        {
            double G = dQ6dQt6(Dp6, Qt6, N6) / 0.595 / 0.995;
            return G;
        }

        public double dNsntodN6dN6()
        {
            double sn = -0.0000211825 * 2 + 0.0048747558;
            return sn;
        }
        public double dNsnpndN6dN6()
        {
            double npn = 0.0085088 * (dQ6dN6dN6(Dp6, Qt6) / 0.595 / 0.995);
            return npn;
        }
        public double dGpvdN6dN6()
        {
            double G = dQ6dN6dN6(Dp6, Qt6) / 0.595 / 0.995;
            return G;
        }
        public double dNsnpndQt6dQt6()
        {
            double npn = 0.0085088 * dQ6dQt6dQt6(Dp6, N6) / 0.595 / 0.995;
            return npn;
        }
        public double dNsnpndQt6dN6()
        {
            double npn = 0.0085088 * dQ6dQt6dN6(Dp6, Qt6, N6) / 0.595 / 0.995;
            return npn;
        }
        public double dGpvdQt6dQt6()
        {
            double G = dQ6dQt6dQt6(Dp6, N6) / 0.595 / 0.995;
            return G;
        }
        public double dNsnpndDp6dQt6()
        {
            double npn = 0.0085088 * (dQ6dDp6dQt6(Dp6, Qt6, N6) / 0.595 / 0.995);
            return npn;
        }
        public double dNsnpndDp6dN6()
        {
            double npn = 0.0085088 * (dQ6dDp6dN6(Dp6, Qt6, N6) / 0.595 / 0.995);
            return npn;
        }
        public double dNsnpndDp6dDp6()
        {
            double npn = 0.0085088 * (dQ6dDp6dDp6(Qt6, N6) / 0.595 / 0.995);
            return npn;
        }
        public double dGpvdQt6dN6()
        {
            double G = dQ6dQt6dN6(Dp6, Qt6, N6) / 0.595 / 0.995;
            return G;
        }

        public double dGpvdDp6dQt6()
        {
            double G = dQ6dDp6dQt6(Dp6, Qt6, N6) / 0.595 / 0.995;
            return G;
        }
        public double dGpvdDp6dN6()
        {
            double G = dQ6dDp6dN6(Dp6, Qt6, N6) / 0.595 / 0.995;
            return G;
        }
        public double dGpvdDp6dDp6()
        {
            double G = dQ6dDp6dDp6(Qt6, N6) / 0.595 / 0.995;
            return G;
        }

        public double dNsntodN4dN4()
        {
            double sn = -0.0000134196 * 2;
            return sn;
        }

        public double dNsnpndQt4dN4()
        {
            double npn = 0.0085088 * dQ4dN4dQt4(P4, Qt4, N4) / 0.595 / 0.995;
            return npn;
        }
        public double dNsnpndQt4dQt4()
        {
            double npn = 0.0085088 * dQ4dQt4dQt4(P4, N4) / 0.595 / 0.995;
            return npn;
        }
        public double dNsnpndN4dN4()
        {
            double npn = 0.0085088 * (dQ4dN4dN4(P4, Qt4) / 0.595 / 0.995);
            return npn;
        }
        public double dGpvdQt4dN4()
        {
            double G = dQ4dN4dQt4(P4, Qt4, N4) / 0.595 / 0.995;
            return G;
        }
        public double dGpvdQt4dQt4()
        {
            double G = dQ4dQt4dQt4(P4, N4) / 0.595 / 0.995;
            return G;
        }
        public double dGpvdN4dN4()
        {
            double G = dQ4dN4dN4(P4, Qt4) / 0.595 / 0.995;
            return G;
        }

        public double dLpdN6()
        {
            double F = lp * (dDauxdN6() - dDQ6dN6(Dp6, Qt6, N6) / (h13 - hv));
            return F;
        }
        public double dLpdN6dN6()
        {
            double F = lp * (dDauxdN6dN6() - dDQ6dN6dN6(Dp6, Qt6) / (h13 - hv));
            return F;
        }
        public double dLpdQt6()
        {
            double F = lp * (dDauxdQt6() - dDQ6dQt6(Dp6, Qt6, N6) / (h13 - hv));
            return F;
        }
        public double dLpdQt6dN6()
        {
            double F = lp * (dDauxdQt6dN6() - dDQ6dQt6dN6(Dp6, Qt6, N6) / (h13 - hv));
            return F;
        }
        public double dLpdQt6dQt6()
        {
            double F = lp * (dDauxdQt6dQt6() - dDQ6dQt6dQt6(Dp6, N6) / (h13 - hv));
            return F;
        }
        public double dLpdDp6()
        {
            double F = lp * (dDauxdDp6() - (dDQ6dDp6(Dp6, Qt6, N6) - hv) / (h13 - hv));
            return F;
        }
        public double dLpdDp6dN6()
        {
            double F = lp * (dDauxdDp6dN6() - dDQ6dDp6dN6(Dp6, Qt6, N6) / (h13 - hv));
            return F;
        }
        public double dLpdDp6dQt6()
        {
            double F = lp * (dDauxdDp6dQt6() - dDQ6dDp6dQt6(Dp6, Qt6, N6) / (h13 - hv));
            return F;
        }
        public double dLpdDp6dDp6()
        {
            double F = lp * (dDauxdDp6dDp6() - dDQ6dDp6dDp6(Qt6, N6) / (h13 - hv));
            return F;
        }

        public double dLpdN4()
        {
            double F = 0;
            F = lp * dDauxdN4();
            return F;
        }
        public double dLpdN4dN4()
        {
            double F = lp * dDauxdN4dN4();
            return F;
        }
        public double dLpdQt4()
        {
            double F = lp * dDauxdQt4();
            return F;
        }
        public double dLpdQt4dN4()
        {
            double F = lp * dDauxdQt4dN4();
            return F;
        }
        public double dLpdQt4dQt4()
        {
            double F = lp * dDauxdQt4dQt4();
            return F;
        }

        public double dLpdN3()
        {
            double F = 0;
            F = lp * dDauxdN3();
            return F;
        }
        public double dLpdN3dN3()
        {
            double F = lp * dDauxdN3dN3();
            return F;
        }
        public double dLpdQt3()
        {
            double F = 0;
            F = lp * dDauxdQt3();
            return F;
        }
        public double dLpdQt3dQt3()
        {
            double F = lp * dDauxdQt3dQt3();
            return F;
        }
        public double dLpdQt3dN3()
        {
            double F = lp * dDauxdQt3dN3();
            return F;
        }

        public double dLpdDp2()
        {
            double F = lp * (dDauxdDp2() - (689.9719271 - hv) / (h13 - hv));
            return F;
        }

        public double Dh2(double D)
        {
            double h = 0;
            if (D != 0) { h = 685.93 * D + 6849.2; }
            return h;
        }
        public double D2p(double D) // отпуск с впрыском
        {
            double Dp = 0;
            if (D != 0) { Dp = (D * (685.93 - hv) + 6849) / (h13 - hv); }
            return Dp;
        }
        public double DP2(double Dp) // отпуск без впрыска
        {
            double D = 0;
            if (Math.Abs(Dp) > 1) { D = (Dp * (h13 - hv) - 6849) / (685.93 - hv); }
            return D;
        }
        public double D6p(double D) // отпуск с впрыском
        {
            double Dp = 0;
            if (D != 0) Dp = DQ6(D, Qt6, N6);
            return Dp;
        }

        public double Daux() // 13 ата на СН
        {
            double D13, G1, G3, G4, G6, Dm1, Qm2, Qm3, Qm4, Qm5, Qm6;
            G1 = 0; G3 = 0; G4 = 0; G6 = 0; Dm1 = 0; Qm2 = 0; Qm3 = 0; Qm4 = 0; Qm5 = 0; Qm6 = 0;
            if (indtg1 == 1) G1 = Dk1(P1, Qt1, N1); if (indtg3 == 1) G3 = Dk3(P3, Qt3, N3);
            if (indtg4 == 1) G4 = Dk4(P4, Qt4, N4); if (indtg6 == 1) G6 = Dk6(Dp6, Qt6, N6);
            Dm1 = D1(P1, Qt1, N1); Qm2 = Q2(Dp2); Qm3 = Q3(P3, Qt3, N3); Qm4 = Q4(P4, Qt4, N4); Qm5 = Q5(Dp5); Qm6 = Q6(Dp6, Qt6, N6);
            double PV = ((Qm2 + Qm3 + Qm4 + Qm5 + Qm6) / 0.595 + D140_13 + D140_30 + D140_45) / 0.995;
            double t6 = 0;
            if (indP6 == 1) t6 = 0.175 * (Qm6 / 0.6 * 0.85 - Dp6) + 98.333;
            if (indP6 == 2) t6 = 0.15 * (Qm6 / 0.6 * 0.85 - Dp6) + 90.667;
            double t = ((Dm1 * (0.3677 * Dm1 + 78.346) + (Qm3 * (102.507 + 7.084 * P3 + 0.273 * Qm3) * 0.9 + Qm4 * (96.958 + 4.933 * P4 + 0.194 * Qm4) * 0.9 + Qm6 * t6 * 0.85) / 0.6) + Ghob * 104 + Gdist * 115 + (Qm2 + Qm3 + Qm4 + Qm5 + Qm6) / 0.6 * 0.1 * 165 + Qpb / kpb * 140) / (Dm1 + (Qm3 * 0.9 + Qm4 * 0.9 + Qm6 * 0.85) / 0.6 + Ghob + Gdist + (Qm2 + Qm3 + Qm4 + Qm5 + Qm6) / 0.6 * 0.1 + Qpb / kpb);
            Ddvd = (165 - t) * (Dm1 + (Qm3 * 0.9 + Qm4 * 0.9 + Qm6 * 0.85) / 0.6 + Ghob + Gdist + (Qm2 + Qm3 + Qm4 + Qm5 + Qm6) / 0.6 * 0.1 + Qpb / kpb) / 535;
            Ddnd = Ghob * 64 / 536 - Ddvd * 0.01 - Q2(Dp2) * 0.01 - PV * 0.005 * 0.98;
            Dpsv = (tsw - tcwout(G1, G3, G4, G6, Gcw1, Gcw3, Gcw4, Gcw6, tnv)) * Gsw / 536;
            if (Dpsv < 0) Dpsv = 0;
            D13 = Ddvd + Ddnd + Dpsv + Gmh + Gdist * 0.195 + Gpts * (tpodp - tsw) / 540 + G13pr;
            return D13;
        }
        public double dDauxdDp2()
        {
            double D = 16 * dQ2dDp2() / (0.595 * 0.995 * 535);
            return D;
        }
        public double dDauxdN6()
        {
            double G1, G3, G4, G6, A, B, C, D, E, F, G, H, dC, dE, dF, dG, dH, Gt, Gcwt, dGt, dtk;
            double e = -0.5913;
            double f = 598.49;
            G1 = 0; G3 = 0; G4 = 0; G6 = 0;
            if (indtg1 == 1) G1 = Dk1(P1, Qt1, N1);
            if (indtg3 == 1) G3 = Dk3(P3, Qt3, N3);
            if (indtg4 == 1) G4 = Dk4(P4, Qt4, N4);
            if (indtg6 == 1) G6 = Dk6(Dp6, Qt6, N6);
            Gt = G6; dGt = dDk6dN6(Dp6, Qt6); Gcwt = Gcw6; dtk = dtk6;

            A = DSTR41[0] + DSTR41[1] * tnv; B = DSTR41[2] + DSTR41[3] * tnv;
            C = G1 / (1 - e * G1 / Gcw1) + G3 / (1 - e * G3 / Gcw3) + G4 / (1 - e * G4 / Gcw4) + G6 / (1 - e * G6 / Gcw6);
            E = G1 * dtk1 / (1 - e * G1 / Gcw1) + G3 * dtk3 / (1 - e * G3 / Gcw3) + G4 * dtk4 / (1 - e * G4 / Gcw4) + G6 * dtk6 / (1 - e * G6 / Gcw6);
            D = Gcw1 + Gcw3 + Gcw4 + Gcw6;
            F = (A * D + B * e * E + B * f * C) / (D - B * e * C);
            G = D + Gcw1 * G1 * e * (Gcw1 - e * G1) + Gcw3 * G3 * e * (Gcw3 - e * G3) + Gcw4 * G4 * e * (Gcw4 - e * G4) + Gcw6 * G6 * e * (Gcw6 - e * G6);
            H = Gcw1 * G1 * (e * dtk1 + f) / (Gcw1 - e * G1) + Gcw3 * G3 * (e * dtk3 + f) / (Gcw3 - e * G3) + Gcw4 * G4 * (e * dtk4 + f) / (Gcw4 - e * G4) + Gcw6 * G6 * (e * dtk6 + f) / (Gcw6 - e * G6);
            dC = (dGt + Gt * e * (1 - dGt) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2);
            dE = dtk * dC;
            dF = B * ((e * dE + f * dC) * (D - B * e * C) + (A * D + B * e * E + B * f * C) * e * dC) / Math.Pow((D - B * e * C), 2);
            dG = Math.Pow(Gcwt, 2) * e * dGt / Math.Pow((Gcwt - e * Gt), 2);
            dH = (e * dtk + f) * dG / e;
            double faux = Ddvd + 1 / D * (dF * G + F * dG + dH);
            return faux;
        }
        public double dDauxdQt6()
        {
            Ddvd = 16 * dQ6dQt6(Dp6, Qt6, N6) / (0.595 * 0.995 * 535);
            double G1, G3, G4, G6, A, B, C, D, E, F, G, H, dC, dE, dF, dG, dH, Gt, Gcwt, dGt, dtk;
            double e = -0.5913;
            double f = 598.49;
            G1 = 0; G3 = 0; G4 = 0; G6 = 0;
            if (indtg1 == 1) G1 = Dk1(P1, Qt1, N1);
            if (indtg3 == 1) G3 = Dk3(P3, Qt3, N3);
            if (indtg4 == 1) G4 = Dk4(P4, Qt4, N4);
            if (indtg6 == 1) G6 = Dk6(Dp6, Qt6, N6);
            Gt = G6; dGt = dDk6dQt6(Dp6, N6); Gcwt = Gcw6; dtk = dtk6;

            A = DSTR41[0] + DSTR41[1] * tnv; B = DSTR41[2] + DSTR41[3] * tnv;
            C = G1 / (1 - e * G1 / Gcw1) + G3 / (1 - e * G3 / Gcw3) + G4 / (1 - e * G4 / Gcw4) + G6 / (1 - e * G6 / Gcw6);
            E = G1 * dtk1 / (1 - e * G1 / Gcw1) + G3 * dtk3 / (1 - e * G3 / Gcw3) + G4 * dtk4 / (1 - e * G4 / Gcw4) + G6 * dtk6 / (1 - e * G6 / Gcw6);
            D = Gcw1 + Gcw3 + Gcw4 + Gcw6;
            F = (A * D + B * e * E + B * f * C) / (D - B * e * C);
            G = D + Gcw1 * G1 * e * (Gcw1 - e * G1) + Gcw3 * G3 * e * (Gcw3 - e * G3) + Gcw4 * G4 * e * (Gcw4 - e * G4) + Gcw6 * G6 * e * (Gcw6 - e * G6);
            H = Gcw1 * G1 * (e * dtk1 + f) / (Gcw1 - e * G1) + Gcw3 * G3 * (e * dtk3 + f) / (Gcw3 - e * G3) + Gcw4 * G4 * (e * dtk4 + f) / (Gcw4 - e * G4) + Gcw6 * G6 * (e * dtk6 + f) / (Gcw6 - e * G6);
            dC = (dGt + Gt * e * (1 - dGt) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2);
            dE = dtk * dC;
            dF = B * ((e * dE + f * dC) * (D - B * e * C) + (A * D + B * e * E + B * f * C) * e * dC) / Math.Pow((D - B * e * C), 2);
            dG = Math.Pow(Gcwt, 2) * e * dGt / Math.Pow((Gcwt - e * Gt), 2);
            dH = (e * dtk + f) * dG / e;
            double faux = Ddvd + 1 / D * (dF * G + F * dG + dH);
            return faux;
        }
        public double dDauxdDp6()
        {
            Ddvd = 16 * dQ6dDp6(Dp6, Qt6, N6) / (0.595 * 0.995 * 535);
            double G1, G3, G4, G6, A, B, C, D, E, F, G, H, dC, dE, dF, dG, dH, Gt, Gcwt, dGt, dtk;
            double e = -0.5913;
            double f = 598.49;
            G1 = 0; G3 = 0; G4 = 0; G6 = 0;
            if (indtg1 == 1) G1 = Dk1(P1, Qt1, N1);
            if (indtg3 == 1) G3 = Dk3(P3, Qt3, N3);
            if (indtg4 == 1) G4 = Dk4(P4, Qt4, N4);
            if (indtg6 == 1) G6 = Dk6(Dp6, Qt6, N6);
            Gt = G6; dGt = dDk6dDp6(Qt6, N6); Gcwt = Gcw6; dtk = dtk6;

            A = DSTR41[0] + DSTR41[1] * tnv; B = DSTR41[2] + DSTR41[3] * tnv;
            C = G1 / (1 - e * G1 / Gcw1) + G3 / (1 - e * G3 / Gcw3) + G4 / (1 - e * G4 / Gcw4) + G6 / (1 - e * G6 / Gcw6);
            E = G1 * dtk1 / (1 - e * G1 / Gcw1) + G3 * dtk3 / (1 - e * G3 / Gcw3) + G4 * dtk4 / (1 - e * G4 / Gcw4) + G6 * dtk6 / (1 - e * G6 / Gcw6);
            D = Gcw1 + Gcw3 + Gcw4 + Gcw6;
            F = (A * D + B * e * E + B * f * C) / (D - B * e * C);
            G = D + Gcw1 * G1 * e * (Gcw1 - e * G1) + Gcw3 * G3 * e * (Gcw3 - e * G3) + Gcw4 * G4 * e * (Gcw4 - e * G4) + Gcw6 * G6 * e * (Gcw6 - e * G6);
            H = Gcw1 * G1 * (e * dtk1 + f) / (Gcw1 - e * G1) + Gcw3 * G3 * (e * dtk3 + f) / (Gcw3 - e * G3) + Gcw4 * G4 * (e * dtk4 + f) / (Gcw4 - e * G4) + Gcw6 * G6 * (e * dtk6 + f) / (Gcw6 - e * G6);
            dC = (dGt + Gt * e * (1 - dGt) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2);
            dE = dtk * dC;
            dF = B * ((e * dE + f * dC) * (D - B * e * C) + (A * D + B * e * E + B * f * C) * e * dC) / Math.Pow((D - B * e * C), 2);
            dG = Math.Pow(Gcwt, 2) * e * dGt / Math.Pow((Gcwt - e * Gt), 2);
            dH = (e * dtk + f) * dG / e;
            double faux = Ddvd + 1 / D * (dF * G + F * dG + dH);
            return faux;
        }

        public double dDauxdDp6dN6()
        {
            Ddvd = 16 * dQ6dDp6dN6(Dp6, Qt6, N6) / (0.595 * 0.995 * 535);
            double G1, G3, G4, G6, A, B, C, D, E, F, G, H, dC, dE, dF, dG, dH, Gt, Gcwt, dGt, ddGt, dtk, ddC, ddE, ddF, ddG, ddH;
            double e = -0.5913;
            double f = 598.49;
            G1 = 0; G3 = 0; G4 = 0; G6 = 0;
            if (indtg1 == 1) G1 = Dk1(P1, Qt1, N1);
            if (indtg3 == 1) G3 = Dk3(P3, Qt3, N3);
            if (indtg4 == 1) G4 = Dk4(P4, Qt4, N4);
            if (indtg6 == 1) G6 = Dk6(Dp6, Qt6, N6);
            Gt = G6; dGt = dDk6dDp6(Qt6, N6); Gcwt = Gcw6; dtk = dtk6; ddGt = dDk6dDp6dN6(Qt6);

            A = DSTR41[0] + DSTR41[1] * tnv; B = DSTR41[2] + DSTR41[3] * tnv;
            C = G1 / (1 - e * G1 / Gcw1) + G3 / (1 - e * G3 / Gcw3) + G4 / (1 - e * G4 / Gcw4) + G6 / (1 - e * G6 / Gcw6);
            E = G1 * dtk1 / (1 - e * G1 / Gcw1) + G3 * dtk3 / (1 - e * G3 / Gcw3) + G4 * dtk4 / (1 - e * G4 / Gcw4) + G6 * dtk6 / (1 - e * G6 / Gcw6);
            D = Gcw1 + Gcw3 + Gcw4 + Gcw6;
            F = (A * D + B * e * E + B * f * C) / (D - B * e * C);
            G = D + Gcw1 * G1 * e * (Gcw1 - e * G1) + Gcw3 * G3 * e * (Gcw3 - e * G3) + Gcw4 * G4 * e * (Gcw4 - e * G4) + Gcw6 * G6 * e * (Gcw6 - e * G6);
            H = Gcw1 * G1 * (e * dtk1 + f) / (Gcw1 - e * G1) + Gcw3 * G3 * (e * dtk3 + f) / (Gcw3 - e * G3) + Gcw4 * G4 * (e * dtk4 + f) / (Gcw4 - e * G4) + Gcw6 * G6 * (e * dtk6 + f) / (Gcw6 - e * G6);
            dC = (dGt + Gt * e * (1 - dGt) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2);
            ddC = (ddGt * (1 - e * Gt / Gcwt) + 2 * e * Math.Pow(dGt, 2) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2) + e * ((dGt * (1 - dGt) - Gt * ddGt) * (1 - e * Gt / Gcwt) + Gt * (1 - dGt) * 2 * e * dGt / Gcwt) / (Gcwt * Math.Pow((1 - e * Gt / Gcwt), 3));
            ddE = dtk * ddC;
            dE = dtk * dC;
            dF = B * ((e * dE + f * dC) * (D - B * e * C) + (A * D + B * e * E + B * f * C) * e * dC) / Math.Pow((D - B * e * C), 2);
            dG = Math.Pow(Gcwt, 2) * e * dGt / Math.Pow((Gcwt - e * Gt), 2);
            dH = (e * dtk + f) * dG / e;
            ddF = B * ((((e * ddE + f * ddC) * (D - B * e * C) + (e * dE + f * dC) * B * e * C) / Math.Pow((D - B * e * C), 2)) + (e * (B * (e * dE + f * dC) * dC + (A * D + B * (e * E + f * C)) * ddC) / Math.Pow((D - B * e * C), 4)));
            ddG = Math.Pow(Gcwt, 2) * e * (ddGt * (Gcwt - e * Gt) + 2 * e * Math.Pow(dGt, 2)) / Math.Pow((Gcwt - e * Gt), 3);
            ddH = (e * dtk + f) * ddG / e;
            double faux = Ddvd + 1 / D * (ddF * G + 2 * dF * dG + F * ddG + ddH);
            return faux;
        }
        public double dDauxdDp6dQt6()
        {
            Ddvd = 16 * dQ6dDp6dQt6(Dp6, Qt6, N6) / (0.595 * 0.995 * 535);
            double G1, G3, G4, G6, A, B, C, D, E, F, G, H, dC, dE, dF, dG, dH, Gt, Gcwt, dGt, ddGt, dtk, ddC, ddE, ddF, ddG, ddH;
            double e = -0.5913;
            double f = 598.49;
            G1 = 0; G3 = 0; G4 = 0; G6 = 0;
            if (indtg1 == 1) G1 = Dk1(P1, Qt1, N1);
            if (indtg3 == 1) G3 = Dk3(P3, Qt3, N3);
            if (indtg4 == 1) G4 = Dk4(P4, Qt4, N4);
            if (indtg6 == 1) G6 = Dk6(Dp6, Qt6, N6);
            Gt = G6; dGt = dDk6dDp6(Qt6, N6); Gcwt = Gcw6; dtk = dtk6; ddGt = dDk6dDp6dQt6(N6);

            A = DSTR41[0] + DSTR41[1] * tnv; B = DSTR41[2] + DSTR41[3] * tnv;
            C = G1 / (1 - e * G1 / Gcw1) + G3 / (1 - e * G3 / Gcw3) + G4 / (1 - e * G4 / Gcw4) + G6 / (1 - e * G6 / Gcw6);
            E = G1 * dtk1 / (1 - e * G1 / Gcw1) + G3 * dtk3 / (1 - e * G3 / Gcw3) + G4 * dtk4 / (1 - e * G4 / Gcw4) + G6 * dtk6 / (1 - e * G6 / Gcw6);
            D = Gcw1 + Gcw3 + Gcw4 + Gcw6;
            F = (A * D + B * e * E + B * f * C) / (D - B * e * C);
            G = D + Gcw1 * G1 * e * (Gcw1 - e * G1) + Gcw3 * G3 * e * (Gcw3 - e * G3) + Gcw4 * G4 * e * (Gcw4 - e * G4) + Gcw6 * G6 * e * (Gcw6 - e * G6);
            H = Gcw1 * G1 * (e * dtk1 + f) / (Gcw1 - e * G1) + Gcw3 * G3 * (e * dtk3 + f) / (Gcw3 - e * G3) + Gcw4 * G4 * (e * dtk4 + f) / (Gcw4 - e * G4) + Gcw6 * G6 * (e * dtk6 + f) / (Gcw6 - e * G6);
            dC = (dGt + Gt * e * (1 - dGt) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2);
            ddC = (ddGt * (1 - e * Gt / Gcwt) + 2 * e * Math.Pow(dGt, 2) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2) + e * ((dGt * (1 - dGt) - Gt * ddGt) * (1 - e * Gt / Gcwt) + Gt * (1 - dGt) * 2 * e * dGt / Gcwt) / (Gcwt * Math.Pow((1 - e * Gt / Gcwt), 3));
            ddE = dtk * ddC;
            dE = dtk * dC;
            dF = B * ((e * dE + f * dC) * (D - B * e * C) + (A * D + B * e * E + B * f * C) * e * dC) / Math.Pow((D - B * e * C), 2);
            dG = Math.Pow(Gcwt, 2) * e * dGt / Math.Pow((Gcwt - e * Gt), 2);
            dH = (e * dtk + f) * dG / e;
            ddF = B * ((((e * ddE + f * ddC) * (D - B * e * C) + (e * dE + f * dC) * B * e * C) / Math.Pow((D - B * e * C), 2)) + (e * (B * (e * dE + f * dC) * dC + (A * D + B * (e * E + f * C)) * ddC) / Math.Pow((D - B * e * C), 4)));
            ddG = Math.Pow(Gcwt, 2) * e * (ddGt * (Gcwt - e * Gt) + 2 * e * Math.Pow(dGt, 2)) / Math.Pow((Gcwt - e * Gt), 3);
            ddH = (e * dtk + f) * ddG / e;
            double faux = Ddvd + 1 / D * (ddF * G + 2 * dF * dG + F * ddG + ddH);
            return faux;
        }
        public double dDauxdDp6dDp6()
        {
            Ddvd = 16 * dQ6dDp6dDp6(Qt6, N6) / (0.595 * 0.995 * 535);
            double G1, G3, G4, G6, A, B, C, D, E, F, G, H, dC, dE, dF, dG, dH, Gt, Gcwt, dGt, ddGt, dtk, ddC, ddE, ddF, ddG, ddH;
            double e = -0.5913;
            double f = 598.49;
            G1 = 0; G3 = 0; G4 = 0; G6 = 0;
            if (indtg1 == 1) G1 = Dk1(P1, Qt1, N1);
            if (indtg3 == 1) G3 = Dk3(P3, Qt3, N3);
            if (indtg4 == 1) G4 = Dk4(P4, Qt4, N4);
            if (indtg6 == 1) G6 = Dk6(Dp6, Qt6, N6);
            Gt = G6; dGt = dDk6dDp6(Qt6, N6); Gcwt = Gcw6; dtk = dtk6; ddGt = 0;

            A = DSTR41[0] + DSTR41[1] * tnv; B = DSTR41[2] + DSTR41[3] * tnv;
            C = G1 / (1 - e * G1 / Gcw1) + G3 / (1 - e * G3 / Gcw3) + G4 / (1 - e * G4 / Gcw4) + G6 / (1 - e * G6 / Gcw6);
            E = G1 * dtk1 / (1 - e * G1 / Gcw1) + G3 * dtk3 / (1 - e * G3 / Gcw3) + G4 * dtk4 / (1 - e * G4 / Gcw4) + G6 * dtk6 / (1 - e * G6 / Gcw6);
            D = Gcw1 + Gcw3 + Gcw4 + Gcw6;
            F = (A * D + B * e * E + B * f * C) / (D - B * e * C);
            G = D + Gcw1 * G1 * e * (Gcw1 - e * G1) + Gcw3 * G3 * e * (Gcw3 - e * G3) + Gcw4 * G4 * e * (Gcw4 - e * G4) + Gcw6 * G6 * e * (Gcw6 - e * G6);
            H = Gcw1 * G1 * (e * dtk1 + f) / (Gcw1 - e * G1) + Gcw3 * G3 * (e * dtk3 + f) / (Gcw3 - e * G3) + Gcw4 * G4 * (e * dtk4 + f) / (Gcw4 - e * G4) + Gcw6 * G6 * (e * dtk6 + f) / (Gcw6 - e * G6);
            dC = (dGt + Gt * e * (1 - dGt) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2);
            ddC = (ddGt * (1 - e * Gt / Gcwt) + 2 * e * Math.Pow(dGt, 2) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2) + e * ((dGt * (1 - dGt) - Gt * ddGt) * (1 - e * Gt / Gcwt) + Gt * (1 - dGt) * 2 * e * dGt / Gcwt) / (Gcwt * Math.Pow((1 - e * Gt / Gcwt), 3));
            ddE = dtk * ddC;
            dE = dtk * dC;
            dF = B * ((e * dE + f * dC) * (D - B * e * C) + (A * D + B * e * E + B * f * C) * e * dC) / Math.Pow((D - B * e * C), 2);
            dG = Math.Pow(Gcwt, 2) * e * dGt / Math.Pow((Gcwt - e * Gt), 2);
            dH = (e * dtk + f) * dG / e;
            ddF = B * ((((e * ddE + f * ddC) * (D - B * e * C) + (e * dE + f * dC) * B * e * C) / Math.Pow((D - B * e * C), 2)) + (e * (B * (e * dE + f * dC) * dC + (A * D + B * (e * E + f * C)) * ddC) / Math.Pow((D - B * e * C), 4)));
            ddG = Math.Pow(Gcwt, 2) * e * (ddGt * (Gcwt - e * Gt) + 2 * e * Math.Pow(dGt, 2)) / Math.Pow((Gcwt - e * Gt), 3);
            ddH = (e * dtk + f) * ddG / e;
            double faux = Ddvd + 1 / D * (ddF * G + 2 * dF * dG + F * ddG + ddH);
            return faux;
        }
        public double dDauxdQt6dN6()
        {
            Ddvd = 16 * dQ6dQt6dN6(Dp6, Qt6, N6) / (0.595 * 0.995 * 535);
            double G1, G3, G4, G6, A, B, C, D, E, F, G, H, dC, dE, dF, dG, dH, Gt, Gcwt, dGt, ddGt, dtk, ddC, ddE, ddF, ddG, ddH;
            double e = -0.5913;
            double f = 598.49;
            G1 = 0; G3 = 0; G4 = 0; G6 = 0;
            if (indtg1 == 1) G1 = Dk1(P1, Qt1, N1);
            if (indtg3 == 1) G3 = Dk3(P3, Qt3, N3);
            if (indtg4 == 1) G4 = Dk4(P4, Qt4, N4);
            if (indtg6 == 1) G6 = Dk6(Dp6, Qt6, N6);
            Gt = G6; dGt = dDk6dQt6(Dp6, N6); Gcwt = Gcw6; dtk = dtk6; ddGt = dDk6dQt6dN6(Dp6);

            A = DSTR41[0] + DSTR41[1] * tnv; B = DSTR41[2] + DSTR41[3] * tnv;
            C = G1 / (1 - e * G1 / Gcw1) + G3 / (1 - e * G3 / Gcw3) + G4 / (1 - e * G4 / Gcw4) + G6 / (1 - e * G6 / Gcw6);
            E = G1 * dtk1 / (1 - e * G1 / Gcw1) + G3 * dtk3 / (1 - e * G3 / Gcw3) + G4 * dtk4 / (1 - e * G4 / Gcw4) + G6 * dtk6 / (1 - e * G6 / Gcw6);
            D = Gcw1 + Gcw3 + Gcw4 + Gcw6;
            F = (A * D + B * e * E + B * f * C) / (D - B * e * C);
            G = D + Gcw1 * G1 * e * (Gcw1 - e * G1) + Gcw3 * G3 * e * (Gcw3 - e * G3) + Gcw4 * G4 * e * (Gcw4 - e * G4) + Gcw6 * G6 * e * (Gcw6 - e * G6);
            H = Gcw1 * G1 * (e * dtk1 + f) / (Gcw1 - e * G1) + Gcw3 * G3 * (e * dtk3 + f) / (Gcw3 - e * G3) + Gcw4 * G4 * (e * dtk4 + f) / (Gcw4 - e * G4) + Gcw6 * G6 * (e * dtk6 + f) / (Gcw6 - e * G6);
            dC = (dGt + Gt * e * (1 - dGt) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2);
            ddC = (ddGt * (1 - e * Gt / Gcwt) + 2 * e * Math.Pow(dGt, 2) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2) + e * ((dGt * (1 - dGt) - Gt * ddGt) * (1 - e * Gt / Gcwt) + Gt * (1 - dGt) * 2 * e * dGt / Gcwt) / (Gcwt * Math.Pow((1 - e * Gt / Gcwt), 3));
            ddE = dtk * ddC;
            dE = dtk * dC;
            dF = B * ((e * dE + f * dC) * (D - B * e * C) + (A * D + B * e * E + B * f * C) * e * dC) / Math.Pow((D - B * e * C), 2);
            dG = Math.Pow(Gcwt, 2) * e * dGt / Math.Pow((Gcwt - e * Gt), 2);
            dH = (e * dtk + f) * dG / e;
            ddF = B * ((((e * ddE + f * ddC) * (D - B * e * C) + (e * dE + f * dC) * B * e * C) / Math.Pow((D - B * e * C), 2)) + (e * (B * (e * dE + f * dC) * dC + (A * D + B * (e * E + f * C)) * ddC) / Math.Pow((D - B * e * C), 4)));
            ddG = Math.Pow(Gcwt, 2) * e * (ddGt * (Gcwt - e * Gt) + 2 * e * Math.Pow(dGt, 2)) / Math.Pow((Gcwt - e * Gt), 3);
            ddH = (e * dtk + f) * ddG / e;
            double faux = Ddvd + 1 / D * (ddF * G + 2 * dF * dG + F * ddG + ddH);
            return faux;
        }
        public double dDauxdQt6dQt6()
        {
            Ddvd = 16 * dQ6dQt6dQt6(Dp6, N6) / (0.595 * 0.995 * 535);
            double G1, G3, G4, G6, A, B, C, D, E, F, G, H, dC, dE, dF, dG, dH, Gt, Gcwt, dGt, ddGt, dtk, ddC, ddE, ddF, ddG, ddH;
            double e = -0.5913;
            double f = 598.49;
            G1 = 0; G3 = 0; G4 = 0; G6 = 0;
            if (indtg1 == 1) G1 = Dk1(P1, Qt1, N1);
            if (indtg3 == 1) G3 = Dk3(P3, Qt3, N3);
            if (indtg4 == 1) G4 = Dk4(P4, Qt4, N4);
            if (indtg6 == 1) G6 = Dk6(Dp6, Qt6, N6);
            Gt = G6; dGt = dDk6dQt6(Dp6, N6); Gcwt = Gcw6; dtk = dtk6; ddGt = 0;

            A = DSTR41[0] + DSTR41[1] * tnv; B = DSTR41[2] + DSTR41[3] * tnv;
            C = G1 / (1 - e * G1 / Gcw1) + G3 / (1 - e * G3 / Gcw3) + G4 / (1 - e * G4 / Gcw4) + G6 / (1 - e * G6 / Gcw6);
            E = G1 * dtk1 / (1 - e * G1 / Gcw1) + G3 * dtk3 / (1 - e * G3 / Gcw3) + G4 * dtk4 / (1 - e * G4 / Gcw4) + G6 * dtk6 / (1 - e * G6 / Gcw6);
            D = Gcw1 + Gcw3 + Gcw4 + Gcw6;
            F = (A * D + B * e * E + B * f * C) / (D - B * e * C);
            G = D + Gcw1 * G1 * e * (Gcw1 - e * G1) + Gcw3 * G3 * e * (Gcw3 - e * G3) + Gcw4 * G4 * e * (Gcw4 - e * G4) + Gcw6 * G6 * e * (Gcw6 - e * G6);
            H = Gcw1 * G1 * (e * dtk1 + f) / (Gcw1 - e * G1) + Gcw3 * G3 * (e * dtk3 + f) / (Gcw3 - e * G3) + Gcw4 * G4 * (e * dtk4 + f) / (Gcw4 - e * G4) + Gcw6 * G6 * (e * dtk6 + f) / (Gcw6 - e * G6);
            dC = (dGt + Gt * e * (1 - dGt) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2);
            ddC = (ddGt * (1 - e * Gt / Gcwt) + 2 * e * Math.Pow(dGt, 2) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2) + e * ((dGt * (1 - dGt) - Gt * ddGt) * (1 - e * Gt / Gcwt) + Gt * (1 - dGt) * 2 * e * dGt / Gcwt) / (Gcwt * Math.Pow((1 - e * Gt / Gcwt), 3));
            ddE = dtk * ddC;
            dE = dtk * dC;
            dF = B * ((e * dE + f * dC) * (D - B * e * C) + (A * D + B * e * E + B * f * C) * e * dC) / Math.Pow((D - B * e * C), 2);
            dG = Math.Pow(Gcwt, 2) * e * dGt / Math.Pow((Gcwt - e * Gt), 2);
            dH = (e * dtk + f) * dG / e;
            ddF = B * ((((e * ddE + f * ddC) * (D - B * e * C) + (e * dE + f * dC) * B * e * C) / Math.Pow((D - B * e * C), 2)) + (e * (B * (e * dE + f * dC) * dC + (A * D + B * (e * E + f * C)) * ddC) / Math.Pow((D - B * e * C), 4)));
            ddG = Math.Pow(Gcwt, 2) * e * (ddGt * (Gcwt - e * Gt) + 2 * e * Math.Pow(dGt, 2)) / Math.Pow((Gcwt - e * Gt), 3);
            ddH = (e * dtk + f) * ddG / e;
            double faux = Ddvd + 1 / D * (ddF * G + 2 * dF * dG + F * ddG + ddH);
            return faux;
        }
        public double dDauxdN6dN6()
        {
            Ddvd = 16 * dQ6dN6dN6(Dp6, Qt6) / (0.595 * 0.995 * 535);
            double G1, G3, G4, G6, A, B, C, D, E, F, G, H, dC, dE, dF, dG, dH, Gt, Gcwt, dGt, ddGt, dtk, ddC, ddE, ddF, ddG, ddH;
            double e = -0.5913;
            double f = 598.49;
            G1 = 0; G3 = 0; G4 = 0; G6 = 0;
            if (indtg1 == 1) G1 = Dk1(P1, Qt1, N1);
            if (indtg3 == 1) G3 = Dk3(P3, Qt3, N3);
            if (indtg4 == 1) G4 = Dk4(P4, Qt4, N4);
            if (indtg6 == 1) G6 = Dk6(Dp6, Qt6, N6);
            Gt = G6; dGt = dDk6dN6(Dp6, Qt6); Gcwt = Gcw6; dtk = dtk6; ddGt = 0;

            A = DSTR41[0] + DSTR41[1] * tnv; B = DSTR41[2] + DSTR41[3] * tnv;
            C = G1 / (1 - e * G1 / Gcw1) + G3 / (1 - e * G3 / Gcw3) + G4 / (1 - e * G4 / Gcw4) + G6 / (1 - e * G6 / Gcw6);
            E = G1 * dtk1 / (1 - e * G1 / Gcw1) + G3 * dtk3 / (1 - e * G3 / Gcw3) + G4 * dtk4 / (1 - e * G4 / Gcw4) + G6 * dtk6 / (1 - e * G6 / Gcw6);
            D = Gcw1 + Gcw3 + Gcw4 + Gcw6;
            F = (A * D + B * e * E + B * f * C) / (D - B * e * C);
            G = D + Gcw1 * G1 * e * (Gcw1 - e * G1) + Gcw3 * G3 * e * (Gcw3 - e * G3) + Gcw4 * G4 * e * (Gcw4 - e * G4) + Gcw6 * G6 * e * (Gcw6 - e * G6);
            H = Gcw1 * G1 * (e * dtk1 + f) / (Gcw1 - e * G1) + Gcw3 * G3 * (e * dtk3 + f) / (Gcw3 - e * G3) + Gcw4 * G4 * (e * dtk4 + f) / (Gcw4 - e * G4) + Gcw6 * G6 * (e * dtk6 + f) / (Gcw6 - e * G6);
            dC = (dGt + Gt * e * (1 - dGt) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2);
            ddC = (ddGt * (1 - e * Gt / Gcwt) + 2 * e * Math.Pow(dGt, 2) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2) + e * ((dGt * (1 - dGt) - Gt * ddGt) * (1 - e * Gt / Gcwt) + Gt * (1 - dGt) * 2 * e * dGt / Gcwt) / (Gcwt * Math.Pow((1 - e * Gt / Gcwt), 3));
            ddE = dtk * ddC;
            dE = dtk * dC;
            dF = B * ((e * dE + f * dC) * (D - B * e * C) + (A * D + B * e * E + B * f * C) * e * dC) / Math.Pow((D - B * e * C), 2);
            dG = Math.Pow(Gcwt, 2) * e * dGt / Math.Pow((Gcwt - e * Gt), 2);
            dH = (e * dtk + f) * dG / e;
            ddF = B * ((((e * ddE + f * ddC) * (D - B * e * C) + (e * dE + f * dC) * B * e * C) / Math.Pow((D - B * e * C), 2)) + (e * (B * (e * dE + f * dC) * dC + (A * D + B * (e * E + f * C)) * ddC) / Math.Pow((D - B * e * C), 4)));
            ddG = Math.Pow(Gcwt, 2) * e * (ddGt * (Gcwt - e * Gt) + 2 * e * Math.Pow(dGt, 2)) / Math.Pow((Gcwt - e * Gt), 3);
            ddH = (e * dtk + f) * ddG / e;
            double faux = Ddvd + 1 / D * (ddF * G + 2 * dF * dG + F * ddG + ddH);
            return faux;
        }

        public double dDauxdN4()
        {
            Ddvd = 16 * dQ4dN4(P4, Qt4, N4) / (0.595 * 0.995 * 535);
            double G1, G3, G4, G6, A, B, C, D, E, F, G, H, dC, dE, dF, dG, dH, Gt, Gcwt, dGt, dtk;
            double e = -0.5913;
            double f = 598.49;
            G1 = 0; G3 = 0; G4 = 0; G6 = 0;
            if (indtg1 == 1) G1 = Dk1(P1, Qt1, N1);
            if (indtg3 == 1) G3 = Dk3(P3, Qt3, N3);
            if (indtg4 == 1) G4 = Dk4(P4, Qt4, N4);
            if (indtg6 == 1) G6 = Dk6(Dp6, Qt6, N6);
            Gt = G4; dGt = dDk4dN4(P4, Qt4); Gcwt = Gcw4; dtk = dtk4;

            A = DSTR41[0] + DSTR41[1] * tnv; B = DSTR41[2] + DSTR41[3] * tnv;
            C = G1 / (1 - e * G1 / Gcw1) + G3 / (1 - e * G3 / Gcw3) + G4 / (1 - e * G4 / Gcw4) + G6 / (1 - e * G6 / Gcw6);
            E = G1 * dtk1 / (1 - e * G1 / Gcw1) + G3 * dtk3 / (1 - e * G3 / Gcw3) + G4 * dtk4 / (1 - e * G4 / Gcw4) + G6 * dtk6 / (1 - e * G6 / Gcw6);
            D = Gcw1 + Gcw3 + Gcw4 + Gcw6;
            F = (A * D + B * e * E + B * f * C) / (D - B * e * C);
            G = D + Gcw1 * G1 * e * (Gcw1 - e * G1) + Gcw3 * G3 * e * (Gcw3 - e * G3) + Gcw4 * G4 * e * (Gcw4 - e * G4) + Gcw6 * G6 * e * (Gcw6 - e * G6);
            H = Gcw1 * G1 * (e * dtk1 + f) / (Gcw1 - e * G1) + Gcw3 * G3 * (e * dtk3 + f) / (Gcw3 - e * G3) + Gcw4 * G4 * (e * dtk4 + f) / (Gcw4 - e * G4) + Gcw6 * G6 * (e * dtk6 + f) / (Gcw6 - e * G6);
            dC = (dGt + Gt * e * (1 - dGt) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2);
            dE = dtk * dC;
            dF = B * ((e * dE + f * dC) * (D - B * e * C) + (A * D + B * e * E + B * f * C) * e * dC) / Math.Pow((D - B * e * C), 2);
            dG = Math.Pow(Gcwt, 2) * e * dGt / Math.Pow((Gcwt - e * Gt), 2);
            dH = (e * dtk + f) * dG / e;
            double faux = Ddvd + 1 / D * (dF * G + F * dG + dH);
            return faux;
        }
        public double dDauxdQt4()
        {
            Ddvd = 16 * dQ4dQt4(P4, Qt4, N4) / (0.595 * 0.995 * 535);
            double G1, G3, G4, G6, A, B, C, D, E, F, G, H, dC, dE, dF, dG, dH, Gt, Gcwt, dGt, dtk;
            double e = -0.5913;
            double f = 598.49;
            G1 = 0; G3 = 0; G4 = 0; G6 = 0;
            if (indtg1 == 1) G1 = Dk1(P1, Qt1, N1);
            if (indtg3 == 1) G3 = Dk3(P3, Qt3, N3);
            if (indtg4 == 1) G4 = Dk4(P4, Qt4, N4);
            if (indtg6 == 1) G6 = Dk6(Dp6, Qt6, N6);
            Gt = G4; dGt = dDk4dQt4(P4, N4); Gcwt = Gcw4; dtk = dtk4;

            A = DSTR41[0] + DSTR41[1] * tnv; B = DSTR41[2] + DSTR41[3] * tnv;
            C = G1 / (1 - e * G1 / Gcw1) + G3 / (1 - e * G3 / Gcw3) + G4 / (1 - e * G4 / Gcw4) + G6 / (1 - e * G6 / Gcw6);
            E = G1 * dtk1 / (1 - e * G1 / Gcw1) + G3 * dtk3 / (1 - e * G3 / Gcw3) + G4 * dtk4 / (1 - e * G4 / Gcw4) + G6 * dtk6 / (1 - e * G6 / Gcw6);
            D = Gcw1 + Gcw3 + Gcw4 + Gcw6;
            F = (A * D + B * e * E + B * f * C) / (D - B * e * C);
            G = D + Gcw1 * G1 * e * (Gcw1 - e * G1) + Gcw3 * G3 * e * (Gcw3 - e * G3) + Gcw4 * G4 * e * (Gcw4 - e * G4) + Gcw6 * G6 * e * (Gcw6 - e * G6);
            H = Gcw1 * G1 * (e * dtk1 + f) / (Gcw1 - e * G1) + Gcw3 * G3 * (e * dtk3 + f) / (Gcw3 - e * G3) + Gcw4 * G4 * (e * dtk4 + f) / (Gcw4 - e * G4) + Gcw6 * G6 * (e * dtk6 + f) / (Gcw6 - e * G6);
            dC = (dGt + Gt * e * (1 - dGt) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2);
            dE = dtk * dC;
            dF = B * ((e * dE + f * dC) * (D - B * e * C) + (A * D + B * e * E + B * f * C) * e * dC) / Math.Pow((D - B * e * C), 2);
            dG = Math.Pow(Gcwt, 2) * e * dGt / Math.Pow((Gcwt - e * Gt), 2);
            dH = (e * dtk + f) * dG / e;
            double faux = Ddvd + 1 / D * (dF * G + F * dG + dH);
            return faux;
        }
        public double dDauxdQt4dN4()
        {
            Ddvd = 16 * dQ4dN4dQt4(P4, Qt4, N4) / (0.595 * 0.995 * 535);
            double G1, G3, G4, G6, A, B, C, D, E, F, G, H, dC, dE, dF, dG, dH, Gt, Gcwt, dGt, ddGt, dtk, ddC, ddE, ddF, ddG, ddH;
            double e = -0.5913;
            double f = 598.49;
            G1 = 0; G3 = 0; G4 = 0; G6 = 0;
            if (indtg1 == 1) G1 = Dk1(P1, Qt1, N1);
            if (indtg3 == 1) G3 = Dk3(P3, Qt3, N3);
            if (indtg4 == 1) G4 = Dk4(P4, Qt4, N4);
            if (indtg6 == 1) G6 = Dk6(Dp6, Qt6, N6);
            Gt = G4; dGt = dDk4dQt4(P4, N4); Gcwt = Gcw4; dtk = dtk4; ddGt = dDk4dQt4dN4(P4);

            A = DSTR41[0] + DSTR41[1] * tnv; B = DSTR41[2] + DSTR41[3] * tnv;
            C = G1 / (1 - e * G1 / Gcw1) + G3 / (1 - e * G3 / Gcw3) + G4 / (1 - e * G4 / Gcw4) + G6 / (1 - e * G6 / Gcw6);
            E = G1 * dtk1 / (1 - e * G1 / Gcw1) + G3 * dtk3 / (1 - e * G3 / Gcw3) + G4 * dtk4 / (1 - e * G4 / Gcw4) + G6 * dtk6 / (1 - e * G6 / Gcw6);
            D = Gcw1 + Gcw3 + Gcw4 + Gcw6;
            F = (A * D + B * e * E + B * f * C) / (D - B * e * C);
            G = D + Gcw1 * G1 * e * (Gcw1 - e * G1) + Gcw3 * G3 * e * (Gcw3 - e * G3) + Gcw4 * G4 * e * (Gcw4 - e * G4) + Gcw6 * G6 * e * (Gcw6 - e * G6);
            H = Gcw1 * G1 * (e * dtk1 + f) / (Gcw1 - e * G1) + Gcw3 * G3 * (e * dtk3 + f) / (Gcw3 - e * G3) + Gcw4 * G4 * (e * dtk4 + f) / (Gcw4 - e * G4) + Gcw6 * G6 * (e * dtk6 + f) / (Gcw6 - e * G6);
            dC = (dGt + Gt * e * (1 - dGt) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2);
            ddC = (ddGt * (1 - e * Gt / Gcwt) + 2 * e * Math.Pow(dGt, 2) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2) + e * ((dGt * (1 - dGt) - Gt * ddGt) * (1 - e * Gt / Gcwt) + Gt * (1 - dGt) * 2 * e * dGt / Gcwt) / (Gcwt * Math.Pow((1 - e * Gt / Gcwt), 3));
            ddE = dtk * ddC;
            dE = dtk * dC;
            dF = B * ((e * dE + f * dC) * (D - B * e * C) + (A * D + B * e * E + B * f * C) * e * dC) / Math.Pow((D - B * e * C), 2);
            dG = Math.Pow(Gcwt, 2) * e * dGt / Math.Pow((Gcwt - e * Gt), 2);
            dH = (e * dtk + f) * dG / e;
            ddF = B * ((((e * ddE + f * ddC) * (D - B * e * C) + (e * dE + f * dC) * B * e * C) / Math.Pow((D - B * e * C), 2)) + (e * (B * (e * dE + f * dC) * dC + (A * D + B * (e * E + f * C)) * ddC) / Math.Pow((D - B * e * C), 4)));
            ddG = Math.Pow(Gcwt, 2) * e * (ddGt * (Gcwt - e * Gt) + 2 * e * Math.Pow(dGt, 2)) / Math.Pow((Gcwt - e * Gt), 3);
            ddH = (e * dtk + f) * ddG / e;
            double faux = Ddvd + 1 / D * (ddF * G + 2 * dF * dG + F * ddG + ddH);
            return faux;
        }
        public double dDauxdQt4dQt4()
        {
            Ddvd = 16 * dQ4dQt4dQt4(P4, N4) / (0.595 * 0.995 * 535);
            double G1, G3, G4, G6, A, B, C, D, E, F, G, H, dC, dE, dF, dG, dH, Gt, Gcwt, dGt, ddGt, dtk, ddC, ddE, ddF, ddG, ddH;
            double e = -0.5913;
            double f = 598.49;
            G1 = 0; G3 = 0; G4 = 0; G6 = 0;
            if (indtg1 == 1) G1 = Dk1(P1, Qt1, N1);
            if (indtg3 == 1) G3 = Dk3(P3, Qt3, N3);
            if (indtg4 == 1) G4 = Dk4(P4, Qt4, N4);
            if (indtg6 == 1) G6 = Dk6(Dp6, Qt6, N6);
            Gt = G4; dGt = dDk4dQt4(P4, N4); Gcwt = Gcw4; dtk = dtk4; ddGt = 0;

            A = DSTR41[0] + DSTR41[1] * tnv; B = DSTR41[2] + DSTR41[3] * tnv;
            C = G1 / (1 - e * G1 / Gcw1) + G3 / (1 - e * G3 / Gcw3) + G4 / (1 - e * G4 / Gcw4) + G6 / (1 - e * G6 / Gcw6);
            E = G1 * dtk1 / (1 - e * G1 / Gcw1) + G3 * dtk3 / (1 - e * G3 / Gcw3) + G4 * dtk4 / (1 - e * G4 / Gcw4) + G6 * dtk6 / (1 - e * G6 / Gcw6);
            D = Gcw1 + Gcw3 + Gcw4 + Gcw6;
            F = (A * D + B * e * E + B * f * C) / (D - B * e * C);
            G = D + Gcw1 * G1 * e * (Gcw1 - e * G1) + Gcw3 * G3 * e * (Gcw3 - e * G3) + Gcw4 * G4 * e * (Gcw4 - e * G4) + Gcw6 * G6 * e * (Gcw6 - e * G6);
            H = Gcw1 * G1 * (e * dtk1 + f) / (Gcw1 - e * G1) + Gcw3 * G3 * (e * dtk3 + f) / (Gcw3 - e * G3) + Gcw4 * G4 * (e * dtk4 + f) / (Gcw4 - e * G4) + Gcw6 * G6 * (e * dtk6 + f) / (Gcw6 - e * G6);
            dC = (dGt + Gt * e * (1 - dGt) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2);
            ddC = (ddGt * (1 - e * Gt / Gcwt) + 2 * e * Math.Pow(dGt, 2) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2) + e * ((dGt * (1 - dGt) - Gt * ddGt) * (1 - e * Gt / Gcwt) + Gt * (1 - dGt) * 2 * e * dGt / Gcwt) / (Gcwt * Math.Pow((1 - e * Gt / Gcwt), 3));
            ddE = dtk * ddC;
            dE = dtk * dC;
            dF = B * ((e * dE + f * dC) * (D - B * e * C) + (A * D + B * e * E + B * f * C) * e * dC) / Math.Pow((D - B * e * C), 2);
            dG = Math.Pow(Gcwt, 2) * e * dGt / Math.Pow((Gcwt - e * Gt), 2);
            dH = (e * dtk + f) * dG / e;
            ddF = B * ((((e * ddE + f * ddC) * (D - B * e * C) + (e * dE + f * dC) * B * e * C) / Math.Pow((D - B * e * C), 2)) + (e * (B * (e * dE + f * dC) * dC + (A * D + B * (e * E + f * C)) * ddC) / Math.Pow((D - B * e * C), 4)));
            ddG = Math.Pow(Gcwt, 2) * e * (ddGt * (Gcwt - e * Gt) + 2 * e * Math.Pow(dGt, 2)) / Math.Pow((Gcwt - e * Gt), 3);
            ddH = (e * dtk + f) * ddG / e;
            double faux = Ddvd + 1 / D * (ddF * G + 2 * dF * dG + F * ddG + ddH);
            return faux;
        }
        public double dDauxdN4dN4()
        {
            Ddvd = 16 * dQ4dN4dN4(P4, Qt4) / (0.595 * 0.995 * 535);
            double G1, G3, G4, G6, A, B, C, D, E, F, G, H, dC, dE, dF, dG, dH, Gt, Gcwt, dGt, ddGt, dtk, ddC, ddE, ddF, ddG, ddH;
            double e = -0.5913;
            double f = 598.49;
            G1 = 0; G3 = 0; G4 = 0; G6 = 0;
            if (indtg1 == 1) G1 = Dk1(P1, Qt1, N1);
            if (indtg3 == 1) G3 = Dk3(P3, Qt3, N3);
            if (indtg4 == 1) G4 = Dk4(P4, Qt4, N4);
            if (indtg6 == 1) G6 = Dk6(Dp6, Qt6, N6);
            Gt = G4; dGt = dDk4dN4(P4, Qt4); Gcwt = Gcw4; dtk = dtk4; ddGt = 0;

            A = DSTR41[0] + DSTR41[1] * tnv; B = DSTR41[2] + DSTR41[3] * tnv;
            C = G1 / (1 - e * G1 / Gcw1) + G3 / (1 - e * G3 / Gcw3) + G4 / (1 - e * G4 / Gcw4) + G6 / (1 - e * G6 / Gcw6);
            E = G1 * dtk1 / (1 - e * G1 / Gcw1) + G3 * dtk3 / (1 - e * G3 / Gcw3) + G4 * dtk4 / (1 - e * G4 / Gcw4) + G6 * dtk6 / (1 - e * G6 / Gcw6);
            D = Gcw1 + Gcw3 + Gcw4 + Gcw6;
            F = (A * D + B * e * E + B * f * C) / (D - B * e * C);
            G = D + Gcw1 * G1 * e * (Gcw1 - e * G1) + Gcw3 * G3 * e * (Gcw3 - e * G3) + Gcw4 * G4 * e * (Gcw4 - e * G4) + Gcw6 * G6 * e * (Gcw6 - e * G6);
            H = Gcw1 * G1 * (e * dtk1 + f) / (Gcw1 - e * G1) + Gcw3 * G3 * (e * dtk3 + f) / (Gcw3 - e * G3) + Gcw4 * G4 * (e * dtk4 + f) / (Gcw4 - e * G4) + Gcw6 * G6 * (e * dtk6 + f) / (Gcw6 - e * G6);
            dC = (dGt + Gt * e * (1 - dGt) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2);
            ddC = (ddGt * (1 - e * Gt / Gcwt) + 2 * e * Math.Pow(dGt, 2) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2) + e * ((dGt * (1 - dGt) - Gt * ddGt) * (1 - e * Gt / Gcwt) + Gt * (1 - dGt) * 2 * e * dGt / Gcwt) / (Gcwt * Math.Pow((1 - e * Gt / Gcwt), 3));
            ddE = dtk * ddC;
            dE = dtk * dC;
            dF = B * ((e * dE + f * dC) * (D - B * e * C) + (A * D + B * e * E + B * f * C) * e * dC) / Math.Pow((D - B * e * C), 2);
            dG = Math.Pow(Gcwt, 2) * e * dGt / Math.Pow((Gcwt - e * Gt), 2);
            dH = (e * dtk + f) * dG / e;
            ddF = B * ((((e * ddE + f * ddC) * (D - B * e * C) + (e * dE + f * dC) * B * e * C) / Math.Pow((D - B * e * C), 2)) + (e * (B * (e * dE + f * dC) * dC + (A * D + B * (e * E + f * C)) * ddC) / Math.Pow((D - B * e * C), 4)));
            ddG = Math.Pow(Gcwt, 2) * e * (ddGt * (Gcwt - e * Gt) + 2 * e * Math.Pow(dGt, 2)) / Math.Pow((Gcwt - e * Gt), 3);
            ddH = (e * dtk + f) * ddG / e;
            double faux = Ddvd + 1 / D * (ddF * G + 2 * dF * dG + F * ddG + ddH);
            return faux;
        }

        public double dDauxdN3()
        {
            Ddvd = 16 * dQ3dN3(P3, Qt3, N3) / (0.595 * 0.995 * 535);
            double G1, G3, G4, G6, A, B, C, D, E, F, G, H, dC, dE, dF, dG, dH, Gt, Gcwt, dGt, dtk;
            double e = -0.5913;
            double f = 598.49;
            G1 = 0; G3 = 0; G4 = 0; G6 = 0;
            if (indtg1 == 1) G1 = Dk1(P1, Qt1, N1);
            if (indtg3 == 1) G3 = Dk3(P3, Qt3, N3);
            if (indtg4 == 1) G4 = Dk4(P4, Qt4, N4);
            if (indtg6 == 1) G6 = Dk6(Dp6, Qt6, N6);
            Gt = G3; dGt = dDk3dN3(P3, Qt3); Gcwt = Gcw3; dtk = dtk3;

            A = DSTR41[0] + DSTR41[1] * tnv; B = DSTR41[2] + DSTR41[3] * tnv;
            C = G1 / (1 - e * G1 / Gcw1) + G3 / (1 - e * G3 / Gcw3) + G4 / (1 - e * G4 / Gcw4) + G6 / (1 - e * G6 / Gcw6);
            E = G1 * dtk1 / (1 - e * G1 / Gcw1) + G3 * dtk3 / (1 - e * G3 / Gcw3) + G4 * dtk4 / (1 - e * G4 / Gcw4) + G6 * dtk6 / (1 - e * G6 / Gcw6);
            D = Gcw1 + Gcw3 + Gcw4 + Gcw6;
            F = (A * D + B * e * E + B * f * C) / (D - B * e * C);
            G = D + Gcw1 * G1 * e * (Gcw1 - e * G1) + Gcw3 * G3 * e * (Gcw3 - e * G3) + Gcw4 * G4 * e * (Gcw4 - e * G4) + Gcw6 * G6 * e * (Gcw6 - e * G6);
            H = Gcw1 * G1 * (e * dtk1 + f) / (Gcw1 - e * G1) + Gcw3 * G3 * (e * dtk3 + f) / (Gcw3 - e * G3) + Gcw4 * G4 * (e * dtk4 + f) / (Gcw4 - e * G4) + Gcw6 * G6 * (e * dtk6 + f) / (Gcw6 - e * G6);
            dC = (dGt + Gt * e * (1 - dGt) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2);
            dE = dtk * dC;
            dF = B * ((e * dE + f * dC) * (D - B * e * C) + (A * D + B * e * E + B * f * C) * e * dC) / Math.Pow((D - B * e * C), 2);
            dG = Math.Pow(Gcwt, 2) * e * dGt / Math.Pow((Gcwt - e * Gt), 2);
            dH = (e * dtk + f) * dG / e;
            double faux = Ddvd + 1 / D * (dF * G + F * dG + dH);
            return faux;
        }
        public double dDauxdQt3()
        {
            Ddvd = 16 * dQ3dQt3(P3, Qt3, N3) / (0.595 * 0.995 * 535);
            double G1, G3, G4, G6, A, B, C, D, E, F, G, H, dC, dE, dF, dG, dH, Gt, Gcwt, dGt, dtk;
            double e = -0.5913;
            double f = 598.49;
            G1 = 0; G3 = 0; G4 = 0; G6 = 0;
            if (indtg1 == 1) G1 = Dk1(P1, Qt1, N1);
            if (indtg3 == 1) G3 = Dk3(P3, Qt3, N3);
            if (indtg4 == 1) G4 = Dk4(P4, Qt4, N4);
            if (indtg6 == 1) G6 = Dk6(Dp6, Qt6, N6);
            Gt = G3; dGt = dDk3dQt3(P3, N3); Gcwt = Gcw3; dtk = dtk3;

            A = DSTR41[0] + DSTR41[1] * tnv; B = DSTR41[2] + DSTR41[3] * tnv;
            C = G1 / (1 - e * G1 / Gcw1) + G3 / (1 - e * G3 / Gcw3) + G4 / (1 - e * G4 / Gcw4) + G6 / (1 - e * G6 / Gcw6);
            E = G1 * dtk1 / (1 - e * G1 / Gcw1) + G3 * dtk3 / (1 - e * G3 / Gcw3) + G4 * dtk4 / (1 - e * G4 / Gcw4) + G6 * dtk6 / (1 - e * G6 / Gcw6);
            D = Gcw1 + Gcw3 + Gcw4 + Gcw6;
            F = (A * D + B * e * E + B * f * C) / (D - B * e * C);
            G = D + Gcw1 * G1 * e * (Gcw1 - e * G1) + Gcw3 * G3 * e * (Gcw3 - e * G3) + Gcw4 * G4 * e * (Gcw4 - e * G4) + Gcw6 * G6 * e * (Gcw6 - e * G6);
            H = Gcw1 * G1 * (e * dtk1 + f) / (Gcw1 - e * G1) + Gcw3 * G3 * (e * dtk3 + f) / (Gcw3 - e * G3) + Gcw4 * G4 * (e * dtk4 + f) / (Gcw4 - e * G4) + Gcw6 * G6 * (e * dtk6 + f) / (Gcw6 - e * G6);
            dC = (dGt + Gt * e * (1 - dGt) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2);
            dE = dtk * dC;
            dF = B * ((e * dE + f * dC) * (D - B * e * C) + (A * D + B * e * E + B * f * C) * e * dC) / Math.Pow((D - B * e * C), 2);
            dG = Math.Pow(Gcwt, 2) * e * dGt / Math.Pow((Gcwt - e * Gt), 2);
            dH = (e * dtk + f) * dG / e;
            double faux = Ddvd + 1 / D * (dF * G + F * dG + dH);
            return faux;
        }
        public double dDauxdQt3dN3()
        {
            Ddvd = 16 * dQ3dQt3dN3(P3, Qt3, N3) / (0.595 * 0.995 * 535);
            double G1, G3, G4, G6, A, B, C, D, E, F, G, H, dC, dE, dF, dG, dH, Gt, Gcwt, dGt, ddGt, dtk, ddC, ddE, ddF, ddG, ddH;
            double e = -0.5913;
            double f = 598.49;
            G1 = 0; G3 = 0; G4 = 0; G6 = 0;
            if (indtg1 == 1) G1 = Dk1(P1, Qt1, N1);
            if (indtg3 == 1) G3 = Dk3(P3, Qt3, N3);
            if (indtg4 == 1) G4 = Dk4(P4, Qt4, N4);
            if (indtg6 == 1) G6 = Dk6(Dp6, Qt6, N6);
            Gt = G3; dGt = dDk3dQt3(P3, N3); Gcwt = Gcw3; dtk = dtk3; ddGt = dDk3dQt3dN3(P3);

            A = DSTR41[0] + DSTR41[1] * tnv; B = DSTR41[2] + DSTR41[3] * tnv;
            C = G1 / (1 - e * G1 / Gcw1) + G3 / (1 - e * G3 / Gcw3) + G4 / (1 - e * G4 / Gcw4) + G6 / (1 - e * G6 / Gcw6);
            E = G1 * dtk1 / (1 - e * G1 / Gcw1) + G3 * dtk3 / (1 - e * G3 / Gcw3) + G4 * dtk4 / (1 - e * G4 / Gcw4) + G6 * dtk6 / (1 - e * G6 / Gcw6);
            D = Gcw1 + Gcw3 + Gcw4 + Gcw6;
            F = (A * D + B * e * E + B * f * C) / (D - B * e * C);
            G = D + Gcw1 * G1 * e * (Gcw1 - e * G1) + Gcw3 * G3 * e * (Gcw3 - e * G3) + Gcw4 * G4 * e * (Gcw4 - e * G4) + Gcw6 * G6 * e * (Gcw6 - e * G6);
            H = Gcw1 * G1 * (e * dtk1 + f) / (Gcw1 - e * G1) + Gcw3 * G3 * (e * dtk3 + f) / (Gcw3 - e * G3) + Gcw4 * G4 * (e * dtk4 + f) / (Gcw4 - e * G4) + Gcw6 * G6 * (e * dtk6 + f) / (Gcw6 - e * G6);
            dC = (dGt + Gt * e * (1 - dGt) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2);
            ddC = (ddGt * (1 - e * Gt / Gcwt) + 2 * e * Math.Pow(dGt, 2) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2) + e * ((dGt * (1 - dGt) - Gt * ddGt) * (1 - e * Gt / Gcwt) + Gt * (1 - dGt) * 2 * e * dGt / Gcwt) / (Gcwt * Math.Pow((1 - e * Gt / Gcwt), 3));
            ddE = dtk * ddC;
            dE = dtk * dC;
            dF = B * ((e * dE + f * dC) * (D - B * e * C) + (A * D + B * e * E + B * f * C) * e * dC) / Math.Pow((D - B * e * C), 2);
            dG = Math.Pow(Gcwt, 2) * e * dGt / Math.Pow((Gcwt - e * Gt), 2);
            dH = (e * dtk + f) * dG / e;
            ddF = B * ((((e * ddE + f * ddC) * (D - B * e * C) + (e * dE + f * dC) * B * e * C) / Math.Pow((D - B * e * C), 2)) + (e * (B * (e * dE + f * dC) * dC + (A * D + B * (e * E + f * C)) * ddC) / Math.Pow((D - B * e * C), 4)));
            ddG = Math.Pow(Gcwt, 2) * e * (ddGt * (Gcwt - e * Gt) + 2 * e * Math.Pow(dGt, 2)) / Math.Pow((Gcwt - e * Gt), 3);
            ddH = (e * dtk + f) * ddG / e;
            double faux = Ddvd + 1 / D * (ddF * G + 2 * dF * dG + F * ddG + ddH);
            return faux;
        }
        public double dDauxdQt3dQt3()
        {
            Ddvd = 16 * dQ3dQt3dQt3(P3, N3) / (0.595 * 0.995 * 535);
            double G1, G3, G4, G6, A, B, C, D, E, F, G, H, dC, dE, dF, dG, dH, Gt, Gcwt, dGt, ddGt, dtk, ddC, ddE, ddF, ddG, ddH;
            double e = -0.5913;
            double f = 598.49;
            G1 = 0; G3 = 0; G4 = 0; G6 = 0;
            if (indtg1 == 1) G1 = Dk1(P1, Qt1, N1);
            if (indtg3 == 1) G3 = Dk3(P3, Qt3, N3);
            if (indtg4 == 1) G4 = Dk4(P4, Qt4, N4);
            if (indtg6 == 1) G6 = Dk6(Dp6, Qt6, N6);
            Gt = G3; dGt = dDk3dQt3(P3, N3); Gcwt = Gcw3; dtk = dtk3; ddGt = 0;

            A = DSTR41[0] + DSTR41[1] * tnv; B = DSTR41[2] + DSTR41[3] * tnv;
            C = G1 / (1 - e * G1 / Gcw1) + G3 / (1 - e * G3 / Gcw3) + G4 / (1 - e * G4 / Gcw4) + G6 / (1 - e * G6 / Gcw6);
            E = G1 * dtk1 / (1 - e * G1 / Gcw1) + G3 * dtk3 / (1 - e * G3 / Gcw3) + G4 * dtk4 / (1 - e * G4 / Gcw4) + G6 * dtk6 / (1 - e * G6 / Gcw6);
            D = Gcw1 + Gcw3 + Gcw4 + Gcw6;
            F = (A * D + B * e * E + B * f * C) / (D - B * e * C);
            G = D + Gcw1 * G1 * e * (Gcw1 - e * G1) + Gcw3 * G3 * e * (Gcw3 - e * G3) + Gcw4 * G4 * e * (Gcw4 - e * G4) + Gcw6 * G6 * e * (Gcw6 - e * G6);
            H = Gcw1 * G1 * (e * dtk1 + f) / (Gcw1 - e * G1) + Gcw3 * G3 * (e * dtk3 + f) / (Gcw3 - e * G3) + Gcw4 * G4 * (e * dtk4 + f) / (Gcw4 - e * G4) + Gcw6 * G6 * (e * dtk6 + f) / (Gcw6 - e * G6);
            dC = (dGt + Gt * e * (1 - dGt) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2);
            ddC = (ddGt * (1 - e * Gt / Gcwt) + 2 * e * Math.Pow(dGt, 2) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2) + e * ((dGt * (1 - dGt) - Gt * ddGt) * (1 - e * Gt / Gcwt) + Gt * (1 - dGt) * 2 * e * dGt / Gcwt) / (Gcwt * Math.Pow((1 - e * Gt / Gcwt), 3));
            ddE = dtk * ddC;
            dE = dtk * dC;
            dF = B * ((e * dE + f * dC) * (D - B * e * C) + (A * D + B * e * E + B * f * C) * e * dC) / Math.Pow((D - B * e * C), 2);
            dG = Math.Pow(Gcwt, 2) * e * dGt / Math.Pow((Gcwt - e * Gt), 2);
            dH = (e * dtk + f) * dG / e;
            ddF = B * ((((e * ddE + f * ddC) * (D - B * e * C) + (e * dE + f * dC) * B * e * C) / Math.Pow((D - B * e * C), 2)) + (e * (B * (e * dE + f * dC) * dC + (A * D + B * (e * E + f * C)) * ddC) / Math.Pow((D - B * e * C), 4)));
            ddG = Math.Pow(Gcwt, 2) * e * (ddGt * (Gcwt - e * Gt) + 2 * e * Math.Pow(dGt, 2)) / Math.Pow((Gcwt - e * Gt), 3);
            ddH = (e * dtk + f) * ddG / e;
            double faux = Ddvd + 1 / D * (ddF * G + 2 * dF * dG + F * ddG + ddH);
            return faux;
        }
        public double dDauxdN3dN3()
        {
            Ddvd = 16 * dQ3dNdN3(P3, Qt3) / (0.595 * 0.995 * 535);
            double G1, G3, G4, G6, A, B, C, D, E, F, G, H, dC, dE, dF, dG, dH, Gt, Gcwt, dGt, ddGt, dtk, ddC, ddE, ddF, ddG, ddH;
            double e = -0.5913;
            double f = 598.49;
            G1 = 0; G3 = 0; G4 = 0; G6 = 0;
            if (indtg1 == 1) G1 = Dk1(P1, Qt1, N1);
            if (indtg3 == 1) G3 = Dk3(P3, Qt3, N3);
            if (indtg4 == 1) G4 = Dk4(P4, Qt4, N4);
            if (indtg6 == 1) G6 = Dk6(Dp6, Qt6, N6);
            Gt = G3; dGt = dDk3dN3(P3, Qt3); Gcwt = Gcw3; dtk = dtk3; ddGt = 0;

            A = DSTR41[0] + DSTR41[1] * tnv; B = DSTR41[2] + DSTR41[3] * tnv;
            C = G1 / (1 - e * G1 / Gcw1) + G3 / (1 - e * G3 / Gcw3) + G4 / (1 - e * G4 / Gcw4) + G6 / (1 - e * G6 / Gcw6);
            E = G1 * dtk1 / (1 - e * G1 / Gcw1) + G3 * dtk3 / (1 - e * G3 / Gcw3) + G4 * dtk4 / (1 - e * G4 / Gcw4) + G6 * dtk6 / (1 - e * G6 / Gcw6);
            D = Gcw1 + Gcw3 + Gcw4 + Gcw6;
            F = (A * D + B * e * E + B * f * C) / (D - B * e * C);
            G = D + Gcw1 * G1 * e * (Gcw1 - e * G1) + Gcw3 * G3 * e * (Gcw3 - e * G3) + Gcw4 * G4 * e * (Gcw4 - e * G4) + Gcw6 * G6 * e * (Gcw6 - e * G6);
            H = Gcw1 * G1 * (e * dtk1 + f) / (Gcw1 - e * G1) + Gcw3 * G3 * (e * dtk3 + f) / (Gcw3 - e * G3) + Gcw4 * G4 * (e * dtk4 + f) / (Gcw4 - e * G4) + Gcw6 * G6 * (e * dtk6 + f) / (Gcw6 - e * G6);
            dC = (dGt + Gt * e * (1 - dGt) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2);
            ddC = (ddGt * (1 - e * Gt / Gcwt) + 2 * e * Math.Pow(dGt, 2) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2) + e * ((dGt * (1 - dGt) - Gt * ddGt) * (1 - e * Gt / Gcwt) + Gt * (1 - dGt) * 2 * e * dGt / Gcwt) / (Gcwt * Math.Pow((1 - e * Gt / Gcwt), 3));
            ddE = dtk * ddC;
            dE = dtk * dC;
            dF = B * ((e * dE + f * dC) * (D - B * e * C) + (A * D + B * e * E + B * f * C) * e * dC) / Math.Pow((D - B * e * C), 2);
            dG = Math.Pow(Gcwt, 2) * e * dGt / Math.Pow((Gcwt - e * Gt), 2);
            dH = (e * dtk + f) * dG / e;
            ddF = B * ((((e * ddE + f * ddC) * (D - B * e * C) + (e * dE + f * dC) * B * e * C) / Math.Pow((D - B * e * C), 2)) + (e * (B * (e * dE + f * dC) * dC + (A * D + B * (e * E + f * C)) * ddC) / Math.Pow((D - B * e * C), 4)));
            ddG = Math.Pow(Gcwt, 2) * e * (ddGt * (Gcwt - e * Gt) + 2 * e * Math.Pow(dGt, 2)) / Math.Pow((Gcwt - e * Gt), 3);
            ddH = (e * dtk + f) * ddG / e;
            double faux = Ddvd + 1 / D * (ddF * G + 2 * dF * dG + F * ddG + ddH);
            return faux;
        }

        public double dDauxdN1()
        {
            double G1, G3, G4, G6, A, B, C, D, E, F, G, H, dC, dE, dF, dG, dH, Gt, Gcwt, dGt, dtk;
            double e = -0.5913;
            double f = 598.49;
            G1 = 0; G3 = 0; G4 = 0; G6 = 0;
            if (indtg1 == 1) G1 = Dk1(P1, Qt1, N1);
            if (indtg3 == 1) G3 = Dk3(P3, Qt3, N3);
            if (indtg4 == 1) G4 = Dk4(P4, Qt4, N4);
            if (indtg6 == 1) G6 = Dk6(Dp6, Qt6, N6);
            Gt = G1; dGt = dDk1dN1(P1, Qt1); Gcwt = Gcw1; dtk = dtk1;

            A = DSTR41[0] + DSTR41[1] * tnv; B = DSTR41[2] + DSTR41[3] * tnv;
            C = G1 / (1 - e * G1 / Gcw1) + G3 / (1 - e * G3 / Gcw3) + G4 / (1 - e * G4 / Gcw4) + G6 / (1 - e * G6 / Gcw6);
            E = G1 * dtk1 / (1 - e * G1 / Gcw1) + G3 * dtk3 / (1 - e * G3 / Gcw3) + G4 * dtk4 / (1 - e * G4 / Gcw4) + G6 * dtk6 / (1 - e * G6 / Gcw6);
            D = Gcw1 + Gcw3 + Gcw4 + Gcw6;
            F = (A * D + B * e * E + B * f * C) / (D - B * e * C);
            G = D + Gcw1 * G1 * e * (Gcw1 - e * G1) + Gcw3 * G3 * e * (Gcw3 - e * G3) + Gcw4 * G4 * e * (Gcw4 - e * G4) + Gcw6 * G6 * e * (Gcw6 - e * G6);
            H = Gcw1 * G1 * (e * dtk1 + f) / (Gcw1 - e * G1) + Gcw3 * G3 * (e * dtk3 + f) / (Gcw3 - e * G3) + Gcw4 * G4 * (e * dtk4 + f) / (Gcw4 - e * G4) + Gcw6 * G6 * (e * dtk6 + f) / (Gcw6 - e * G6);
            dC = (dGt + Gt * e * (1 - dGt) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2);
            dE = dtk * dC;
            dF = B * ((e * dE + f * dC) * (D - B * e * C) + (A * D + B * e * E + B * f * C) * e * dC) / Math.Pow((D - B * e * C), 2);
            dG = Math.Pow(Gcwt, 2) * e * dGt / Math.Pow((Gcwt - e * Gt), 2);
            dH = (e * dtk + f) * dG / e;
            double faux = 1 / D * (dF * G + F * dG + dH);
            return faux;
        }
        public double dDauxdN1dN1()
        {
            double G1, G3, G4, G6, A, B, C, D, E, F, G, H, dC, dE, dF, dG, dH, Gt, Gcwt, dGt, ddGt, dtk, ddC, ddE, ddF, ddG, ddH;
            double e = -0.5913;
            double f = 598.49;
            G1 = 0; G3 = 0; G4 = 0; G6 = 0;
            if (indtg1 == 1) G1 = Dk1(P1, Qt1, N1);
            if (indtg3 == 1) G3 = Dk3(P3, Qt3, N3);
            if (indtg4 == 1) G4 = Dk4(P4, Qt4, N4);
            if (indtg6 == 1) G6 = Dk6(Dp6, Qt6, N6);
            Gt = G1; dGt = dDk1dN1(P1, Qt1); Gcwt = Gcw1; dtk = dtk1; ddGt = 0;

            A = DSTR41[0] + DSTR41[1] * tnv; B = DSTR41[2] + DSTR41[3] * tnv;
            C = G1 / (1 - e * G1 / Gcw1) + G3 / (1 - e * G3 / Gcw3) + G4 / (1 - e * G4 / Gcw4) + G6 / (1 - e * G6 / Gcw6);
            E = G1 * dtk1 / (1 - e * G1 / Gcw1) + G3 * dtk3 / (1 - e * G3 / Gcw3) + G4 * dtk4 / (1 - e * G4 / Gcw4) + G6 * dtk6 / (1 - e * G6 / Gcw6);
            D = Gcw1 + Gcw3 + Gcw4 + Gcw6;
            F = (A * D + B * e * E + B * f * C) / (D - B * e * C);
            G = D + Gcw1 * G1 * e * (Gcw1 - e * G1) + Gcw3 * G3 * e * (Gcw3 - e * G3) + Gcw4 * G4 * e * (Gcw4 - e * G4) + Gcw6 * G6 * e * (Gcw6 - e * G6);
            H = Gcw1 * G1 * (e * dtk1 + f) / (Gcw1 - e * G1) + Gcw3 * G3 * (e * dtk3 + f) / (Gcw3 - e * G3) + Gcw4 * G4 * (e * dtk4 + f) / (Gcw4 - e * G4) + Gcw6 * G6 * (e * dtk6 + f) / (Gcw6 - e * G6);
            dC = (dGt + Gt * e * (1 - dGt) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2);
            ddC = (ddGt * (1 - e * Gt / Gcwt) + 2 * e * Math.Pow(dGt, 2) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2) + e * ((dGt * (1 - dGt) - Gt * ddGt) * (1 - e * Gt / Gcwt) + Gt * (1 - dGt) * 2 * e * dGt / Gcwt) / (Gcwt * Math.Pow((1 - e * Gt / Gcwt), 3));
            ddE = dtk * ddC;
            dE = dtk * dC;
            dF = B * ((e * dE + f * dC) * (D - B * e * C) + (A * D + B * e * E + B * f * C) * e * dC) / Math.Pow((D - B * e * C), 2);
            dG = Math.Pow(Gcwt, 2) * e * dGt / Math.Pow((Gcwt - e * Gt), 2);
            dH = (e * dtk + f) * dG / e;
            ddF = B * ((((e * ddE + f * ddC) * (D - B * e * C) + (e * dE + f * dC) * B * e * C) / Math.Pow((D - B * e * C), 2)) + (e * (B * (e * dE + f * dC) * dC + (A * D + B * (e * E + f * C)) * ddC) / Math.Pow((D - B * e * C), 4)));
            ddG = Math.Pow(Gcwt, 2) * e * (ddGt * (Gcwt - e * Gt) + 2 * e * Math.Pow(dGt, 2)) / Math.Pow((Gcwt - e * Gt), 3);
            ddH = (e * dtk + f) * ddG / e;
            double faux = 1 / D * (ddF * G + 2 * dF * dG + F * ddG + ddH);
            return faux;
        }
        public double dDauxdN1dQt1()
        {
            double G1, G3, G4, G6, A, B, C, D, E, F, G, H, dC, dE, dF, dG, dH, Gt, Gcwt, dGt, ddGt, dtk, ddC, ddE, ddF, ddG, ddH;
            double e = -0.5913;
            double f = 598.49;
            G1 = 0; G3 = 0; G4 = 0; G6 = 0;
            if (indtg1 == 1) G1 = Dk1(P1, Qt1, N1);
            if (indtg3 == 1) G3 = Dk3(P3, Qt3, N3);
            if (indtg4 == 1) G4 = Dk4(P4, Qt4, N4);
            if (indtg6 == 1) G6 = Dk6(Dp6, Qt6, N6);
            Gt = G1; dGt = dDk1dN1(P1, Qt1); Gcwt = Gcw1; dtk = dtk1; ddGt = dDk1dN1dQt1(P1);

            A = DSTR41[0] + DSTR41[1] * tnv; B = DSTR41[2] + DSTR41[3] * tnv;
            C = G1 / (1 - e * G1 / Gcw1) + G3 / (1 - e * G3 / Gcw3) + G4 / (1 - e * G4 / Gcw4) + G6 / (1 - e * G6 / Gcw6);
            E = G1 * dtk1 / (1 - e * G1 / Gcw1) + G3 * dtk3 / (1 - e * G3 / Gcw3) + G4 * dtk4 / (1 - e * G4 / Gcw4) + G6 * dtk6 / (1 - e * G6 / Gcw6);
            D = Gcw1 + Gcw3 + Gcw4 + Gcw6;
            F = (A * D + B * e * E + B * f * C) / (D - B * e * C);
            G = D + Gcw1 * G1 * e * (Gcw1 - e * G1) + Gcw3 * G3 * e * (Gcw3 - e * G3) + Gcw4 * G4 * e * (Gcw4 - e * G4) + Gcw6 * G6 * e * (Gcw6 - e * G6);
            H = Gcw1 * G1 * (e * dtk1 + f) / (Gcw1 - e * G1) + Gcw3 * G3 * (e * dtk3 + f) / (Gcw3 - e * G3) + Gcw4 * G4 * (e * dtk4 + f) / (Gcw4 - e * G4) + Gcw6 * G6 * (e * dtk6 + f) / (Gcw6 - e * G6);
            dC = (dGt + Gt * e * (1 - dGt) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2);
            ddC = (ddGt * (1 - e * Gt / Gcwt) + 2 * e * Math.Pow(dGt, 2) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2) + e * ((dGt * (1 - dGt) - Gt * ddGt) * (1 - e * Gt / Gcwt) + Gt * (1 - dGt) * 2 * e * dGt / Gcwt) / (Gcwt * Math.Pow((1 - e * Gt / Gcwt), 3));
            ddE = dtk * ddC;
            dE = dtk * dC;
            dF = B * ((e * dE + f * dC) * (D - B * e * C) + (A * D + B * e * E + B * f * C) * e * dC) / Math.Pow((D - B * e * C), 2);
            dG = Math.Pow(Gcwt, 2) * e * dGt / Math.Pow((Gcwt - e * Gt), 2);
            dH = (e * dtk + f) * dG / e;
            ddF = B * ((((e * ddE + f * ddC) * (D - B * e * C) + (e * dE + f * dC) * B * e * C) / Math.Pow((D - B * e * C), 2)) + (e * (B * (e * dE + f * dC) * dC + (A * D + B * (e * E + f * C)) * ddC) / Math.Pow((D - B * e * C), 4)));
            ddG = Math.Pow(Gcwt, 2) * e * (ddGt * (Gcwt - e * Gt) + 2 * e * Math.Pow(dGt, 2)) / Math.Pow((Gcwt - e * Gt), 3);
            ddH = (e * dtk + f) * ddG / e;
            double faux = 1 / D * (ddF * G + 2 * dF * dG + F * ddG + ddH);
            return faux;
        }
        public double dDauxdQt1()
        {
            double G1, G3, G4, G6, A, B, C, D, E, F, G, H, dC, dE, dF, dG, dH, Gt, Gcwt, dGt, dtk;
            double e = -0.5913;
            double f = 598.49;
            G1 = 0; G3 = 0; G4 = 0; G6 = 0;
            if (indtg1 == 1) G1 = Dk1(P1, Qt1, N1);
            if (indtg3 == 1) G3 = Dk3(P3, Qt3, N3);
            if (indtg4 == 1) G4 = Dk4(P4, Qt4, N4);
            if (indtg6 == 1) G6 = Dk6(Dp6, Qt6, N6);
            Gt = G1; dGt = dDk1dQt1(P1, N1); Gcwt = Gcw1; dtk = dtk1;

            A = DSTR41[0] + DSTR41[1] * tnv; B = DSTR41[2] + DSTR41[3] * tnv;
            C = G1 / (1 - e * G1 / Gcw1) + G3 / (1 - e * G3 / Gcw3) + G4 / (1 - e * G4 / Gcw4) + G6 / (1 - e * G6 / Gcw6);
            E = G1 * dtk1 / (1 - e * G1 / Gcw1) + G3 * dtk3 / (1 - e * G3 / Gcw3) + G4 * dtk4 / (1 - e * G4 / Gcw4) + G6 * dtk6 / (1 - e * G6 / Gcw6);
            D = Gcw1 + Gcw3 + Gcw4 + Gcw6;
            F = (A * D + B * e * E + B * f * C) / (D - B * e * C);
            G = D + Gcw1 * G1 * e * (Gcw1 - e * G1) + Gcw3 * G3 * e * (Gcw3 - e * G3) + Gcw4 * G4 * e * (Gcw4 - e * G4) + Gcw6 * G6 * e * (Gcw6 - e * G6);
            H = Gcw1 * G1 * (e * dtk1 + f) / (Gcw1 - e * G1) + Gcw3 * G3 * (e * dtk3 + f) / (Gcw3 - e * G3) + Gcw4 * G4 * (e * dtk4 + f) / (Gcw4 - e * G4) + Gcw6 * G6 * (e * dtk6 + f) / (Gcw6 - e * G6);
            dC = (dGt + Gt * e * (1 - dGt) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2);
            dE = dtk * dC;
            dF = B * ((e * dE + f * dC) * (D - B * e * C) + (A * D + B * e * E + B * f * C) * e * dC) / Math.Pow((D - B * e * C), 2);
            dG = Math.Pow(Gcwt, 2) * e * dGt / Math.Pow((Gcwt - e * Gt), 2);
            dH = (e * dtk + f) * dG / e;
            double faux = Ddvd + 1 / D * (dF * G + F * dG + dH);
            return faux;
        }
        public double dDauxdQt1dQt1()
        {
            double G1, G3, G4, G6, A, B, C, D, E, F, G, H, dC, dE, dF, dG, dH, Gt, Gcwt, dGt, ddGt, dtk, ddC, ddE, ddF, ddG, ddH;
            double e = -0.5913;
            double f = 598.49;
            G1 = 0; G3 = 0; G4 = 0; G6 = 0;
            if (indtg1 == 1) G1 = Dk1(P1, Qt1, N1);
            if (indtg3 == 1) G3 = Dk3(P3, Qt3, N3);
            if (indtg4 == 1) G4 = Dk4(P4, Qt4, N4);
            if (indtg6 == 1) G6 = Dk6(Dp6, Qt6, N6);
            Gt = G1; dGt = dDk1dQt1(P1, N1); Gcwt = Gcw1; dtk = dtk1; ddGt = 0;

            A = DSTR41[0] + DSTR41[1] * tnv; B = DSTR41[2] + DSTR41[3] * tnv;
            C = G1 / (1 - e * G1 / Gcw1) + G3 / (1 - e * G3 / Gcw3) + G4 / (1 - e * G4 / Gcw4) + G6 / (1 - e * G6 / Gcw6);
            E = G1 * dtk1 / (1 - e * G1 / Gcw1) + G3 * dtk3 / (1 - e * G3 / Gcw3) + G4 * dtk4 / (1 - e * G4 / Gcw4) + G6 * dtk6 / (1 - e * G6 / Gcw6);
            D = Gcw1 + Gcw3 + Gcw4 + Gcw6;
            F = (A * D + B * e * E + B * f * C) / (D - B * e * C);
            G = D + Gcw1 * G1 * e * (Gcw1 - e * G1) + Gcw3 * G3 * e * (Gcw3 - e * G3) + Gcw4 * G4 * e * (Gcw4 - e * G4) + Gcw6 * G6 * e * (Gcw6 - e * G6);
            H = Gcw1 * G1 * (e * dtk1 + f) / (Gcw1 - e * G1) + Gcw3 * G3 * (e * dtk3 + f) / (Gcw3 - e * G3) + Gcw4 * G4 * (e * dtk4 + f) / (Gcw4 - e * G4) + Gcw6 * G6 * (e * dtk6 + f) / (Gcw6 - e * G6);
            dC = (dGt + Gt * e * (1 - dGt) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2);
            ddC = (ddGt * (1 - e * Gt / Gcwt) + 2 * e * Math.Pow(dGt, 2) / Gcwt) / Math.Pow((1 - e * Gt / Gcwt), 2) + e * ((dGt * (1 - dGt) - Gt * ddGt) * (1 - e * Gt / Gcwt) + Gt * (1 - dGt) * 2 * e * dGt / Gcwt) / (Gcwt * Math.Pow((1 - e * Gt / Gcwt), 3));
            ddE = dtk * ddC;
            dE = dtk * dC;
            dF = B * ((e * dE + f * dC) * (D - B * e * C) + (A * D + B * e * E + B * f * C) * e * dC) / Math.Pow((D - B * e * C), 2);
            dG = Math.Pow(Gcwt, 2) * e * dGt / Math.Pow((Gcwt - e * Gt), 2);
            dH = (e * dtk + f) * dG / e;
            ddF = B * ((((e * ddE + f * ddC) * (D - B * e * C) + (e * dE + f * dC) * B * e * C) / Math.Pow((D - B * e * C), 2)) + (e * (B * (e * dE + f * dC) * dC + (A * D + B * (e * E + f * C)) * ddC) / Math.Pow((D - B * e * C), 4)));
            ddG = Math.Pow(Gcwt, 2) * e * (ddGt * (Gcwt - e * Gt) + 2 * e * Math.Pow(dGt, 2)) / Math.Pow((Gcwt - e * Gt), 3);
            ddH = (e * dtk + f) * ddG / e;
            double faux = 1 / D * (ddF * G + 2 * dF * dG + F * ddG + ddH);
            return faux;
        }

        public double Pk3(double Gk, double t)
        {
            double TGk = 0;
            if (indtg3 == 1)
            {
                double[] X01 = { 1, Gk, t, t * Gk, Math.Pow(Gk, 2), Math.Pow(t, 2), Math.Pow(Gk, 2) * t, Gk * Math.Pow(t, 2), Math.Pow(Gk, 2) * Math.Pow(t, 2) };
                if (Gcw3 == 3500)
                {
                    for (int i = 0; i < 9; i++)
                    {
                        TGk = TGk + X01[i] * DSTR32[i];
                    }
                }
                if (Gcw3 == 7000)
                {
                    for (int i = 0; i < 9; i++)
                    {
                        TGk = TGk + X01[i] * DSTR33[i];
                    }
                }
            }
            return TGk;
        }
        public double dPk3(double Gk, double t)
        {
            double TGk = 0;
            double[] X01 = { 0, 1, 0, t, 2 * Gk, 0, 2 * Gk * t, Math.Pow(t, 2), 2 * Gk * Math.Pow(t, 2) };
            if (Gcw3 == 3500)
            {
                for (int i = 0; i < 9; i++)
                {
                    TGk = TGk + X01[i] * DSTR32[i];
                }
            }
            if (Gcw3 == 7000)
            {
                for (int i = 0; i < 9; i++)
                {
                    TGk = TGk + X01[i] * DSTR33[i];
                }
            }
            return TGk;
        }
        public double ddPk3(double t)
        {
            double TGk = 0;
            double[] X01 = { 0, 0, 0, 0, 2, 0, 2 * t, 0, 2 * Math.Pow(t, 2) };
            if (Gcw3 == 3500)
            {
                for (int i = 0; i < 9; i++)
                {
                    TGk = TGk + X01[i] * DSTR32[i];
                }
            }
            if (Gcw3 == 7000)
            {
                for (int i = 0; i < 9; i++)
                {
                    TGk = TGk + X01[i] * DSTR33[i];
                }
            }
            return TGk;
        }

        public double Pk4(double Gk, double t)
        {
            double TGk = 0;
            if (indtg4 == 1)
            {
                double[] X01 = { 1, Gk, t, t * Gk, Math.Pow(Gk, 2), Math.Pow(t, 2), Math.Pow(Gk, 2) * t, Gk * Math.Pow(t, 2), Math.Pow(Gk, 2) * Math.Pow(t, 2) };
                if (Gcw4 == 8000)
                {
                    for (int i = 0; i < 9; i++)
                    {
                        TGk = TGk + X01[i] * DSTR34[i];
                    }
                }
                if (Gcw4 == 16000)
                {
                    for (int i = 0; i < 9; i++)
                    {
                        TGk = TGk + X01[i] * DSTR35[i];
                    }
                }
            }
            return TGk;
        }
        public double dPk4(double Gk, double t)
        {
            double TGk = 0;
            double[] X01 = { 0, 1, 0, t, 2 * Gk, 0, 2 * Gk * t, Math.Pow(t, 2), 2 * Gk * Math.Pow(t, 2) };
            if (Gcw4 == 8000)
            {
                for (int i = 0; i < 9; i++)
                {
                    TGk = TGk + X01[i] * DSTR34[i];
                }
            }
            if (Gcw4 == 16000)
            {
                for (int i = 0; i < 9; i++)
                {
                    TGk = TGk + X01[i] * DSTR35[i];
                }
            }
            return TGk;
        }
        public double ddPk4(double t)
        {
            double TGk = 0;
            double[] X01 = { 0, 0, 0, 0, 2, 0, 2 * t, 0, 2 * Math.Pow(t, 2) };
            if (Gcw4 == 8000)
            {
                for (int i = 0; i < 9; i++)
                {
                    TGk = TGk + X01[i] * DSTR34[i];
                }
            }
            if (Gcw4 == 16000)
            {
                for (int i = 0; i < 9; i++)
                {
                    TGk = TGk + X01[i] * DSTR35[i];
                }
            }
            return TGk;
        }

        public double Pk6(double Gk, double t)
        {
            double TGk = 0;
            if (indtg6 == 1)
            {
                double[] X01 = { 1, Gk, t, t * Gk, Math.Pow(t, 2), Gk * Math.Pow(t, 2) };
                for (int i = 0; i < 6; i++)
                {
                    TGk = TGk + X01[i] * DSTR36[i];
                }
            }
            return TGk;
        }
        public double dPk6(double Gk, double t)
        {
            double TGk = 0;
            double[] X01 = { 0, 1, 0, t, 0, Math.Pow(t, 2) };
            for (int i = 0; i < 6; i++)
            {
                TGk = TGk + X01[i] * DSTR36[i];
            }
            return TGk;
        }

        public double Pn(double t)
        {
            double p = 0.0000023 * Math.Pow(t, 3) - 0.0001887 * Math.Pow(t, 2) + 0.0063300 * t - 0.0208497;
            return p;
        }

        public double Dk1(double P, double Qt, double N)
        {
            double TGk = 0;
            double[] X01 = { 1, P, Qt, N, P * Qt, Qt * N, P * N, P * Qt * N };
            for (int i = 0; i < 8; i++)
            {
                TGk = TGk + X01[i] * DSTR8[i];
            }
            return TGk;
        }
        public double dDk1dN1(double P, double Qt)
        {
            double TGk = 0;
            double[] X01 = { 0, 0, 0, 1, 0, Qt, P, P * Qt };
            for (int i = 0; i < 8; i++)
            {
                TGk = TGk + X01[i] * DSTR8[i];
            }
            return TGk;
        }
        public double dDk1dN1dQt1(double P)
        {
            double TGk = 0;
            double[] X01 = { 0, 0, 0, 0, 0, 1, 0, P };
            for (int i = 0; i < 8; i++)
            {
                TGk = TGk + X01[i] * DSTR8[i];
            }
            return TGk;
        }
        public double dDk1dQt1(double P, double N)
        {
            double TGk = 0;
            double[] X01 = { 0, 0, 1, 0, P, N, 0, P * N };
            for (int i = 0; i < 8; i++)
            {
                TGk = TGk + X01[i] * DSTR8[i];
            }
            return TGk;
        }

        public double Dk3(double P, double Qt, double N)
        {
            double TGk = 0;
            if (indtg3 == 1)
            {
                double[] X03 = { 1, P, Qt, N, P * Qt, Qt * N, P * N, P * Qt * N };
                if (indP3 == 1)
                {
                    for (int i = 0; i < 8; i++)
                    {
                        TGk = TGk + X03[i] * DSTR9[i];
                    }
                }
                else
                {
                    for (int i = 0; i < 8; i++)
                    {
                        TGk = TGk + X03[i] * DSTR10[i];
                    }
                }
            }
            return TGk;
        }
        public double dDk3dN3(double P, double Qt)
        {
            double TGk = 0;
            if (indtg3 == 1)
            {
                double[] X03 = { 0, 0, 0, 1, 0, Qt, P, P * Qt };
                if (indP3 == 1)
                {
                    for (int i = 0; i < 8; i++)
                    {
                        TGk = TGk + X03[i] * DSTR9[i];
                    }
                }
                else
                {
                    for (int i = 0; i < 8; i++)
                    {
                        TGk = TGk + X03[i] * DSTR10[i];
                    }
                }
            }
            return TGk;
        }
        public double dDk3dQt3(double P, double N)
        {
            double TGk = 0;
            if (indtg3 == 1)
            {
                double[] X03 = { 0, 0, 1, 0, P, N, 0, P * N };
                if (indP3 == 1)
                {
                    for (int i = 0; i < 8; i++)
                    {
                        TGk = TGk + X03[i] * DSTR9[i];
                    }
                }
                else
                {
                    for (int i = 0; i < 8; i++)
                    {
                        TGk = TGk + X03[i] * DSTR10[i];
                    }
                }
            }
            return TGk;
        }
        public double dDk3dQt3dN3(double P)
        {
            double TGk = 0;
            if (indtg3 == 1)
            {
                double[] X03 = { 0, 0, 0, 0, 0, 1, 0, P };
                if (indP3 == 1)
                {
                    for (int i = 0; i < 8; i++)
                    {
                        TGk = TGk + X03[i] * DSTR9[i];
                    }
                }
                else
                {
                    for (int i = 0; i < 8; i++)
                    {
                        TGk = TGk + X03[i] * DSTR10[i];
                    }
                }
            }
            return TGk;
        }

        public double Dk4(double P, double Qt, double N)
        {
            double TGk = 0;
            if (indtg4 == 1)
            {
                double[] X04 = { 1, P, Qt, N, P * Qt, Qt * N, P * N, P * Qt * N };
                if (indP4 == 1)
                {
                    for (int i = 0; i < 8; i++)
                    {
                        TGk = TGk + X04[i] * DSTR11[i];
                    }
                }
                else
                {
                    for (int i = 0; i < 8; i++)
                    {
                        TGk = TGk + X04[i] * DSTR12[i];
                    }
                }
            }
            return TGk;
        }
        public double dDk4dN4(double P, double Qt)
        {
            double TGk = 0;
            if (indtg4 == 1)
            {
                double[] X04 = { 0, 0, 0, 1, 0, Qt, P, P * Qt };
                if (indP4 == 1)
                {
                    for (int i = 0; i < 8; i++)
                    {
                        TGk = TGk + X04[i] * DSTR11[i];
                    }
                }
                else
                {
                    for (int i = 0; i < 8; i++)
                    {
                        TGk = TGk + X04[i] * DSTR12[i];
                    }
                }
            }
            return TGk;
        }
        public double dDk4dQt4(double P, double N)
        {
            double TGk = 0;
            if (indtg4 == 1)
            {

                double[] X04 = { 0, 0, 1, 0, P, N, 0, P * N };
                if (indP4 == 1)
                {
                    for (int i = 0; i < 8; i++)
                    {
                        TGk = TGk + X04[i] * DSTR11[i];
                    }
                }
                else
                {
                    for (int i = 0; i < 8; i++)
                    {
                        TGk = TGk + X04[i] * DSTR12[i];
                    }
                }
            }
            return TGk;
        }
        public double dDk4dQt4dN4(double P)
        {
            double TGk = 0;
            if (indtg4 == 1)
            {
                double[] X04 = { 0, 0, 0, 0, 0, 1, 0, P };
                if (indP4 == 1)
                {
                    for (int i = 0; i < 8; i++)
                    {
                        TGk = TGk + X04[i] * DSTR11[i];
                    }
                }
                else
                {
                    for (int i = 0; i < 8; i++)
                    {
                        TGk = TGk + X04[i] * DSTR12[i];
                    }
                }
            }
            return TGk;
        }

        public double Dk6(double D, double Qt, double N)
        {
            double TGk = 0;
            if (indtg6 == 1)
            {
                double[] X06 = { 1, D, Qt, N, D * Qt, Qt * N, D * N, D * Qt * N };
                if (indP6 == 1)
                {
                    for (int i = 0; i < 8; i++)
                    {
                        TGk = TGk + X06[i] * DSTR13[i];
                    }
                }
                else
                {
                    for (int i = 0; i < 8; i++)
                    {
                        TGk = TGk + X06[i] * DSTR14[i];
                    }
                }
            }
            return TGk;
        }
        public double dDk6dN6(double D, double Qt)
        {
            double TGk = 0;
            double[] X06 = { 0, 0, 0, 1, 0, Qt, D, D * Qt };
            if (indP6 == 1)
            {
                for (int i = 0; i < 8; i++)
                {
                    TGk = TGk + X06[i] * DSTR13[i];
                }
            }
            else
            {
                for (int i = 0; i < 8; i++)
                {
                    TGk = TGk + X06[i] * DSTR14[i];
                }
            }
            return TGk;
        }
        public double dDk6dQt6(double D, double N)
        {
            double TGk = 0;
            double[] X06 = { 0, 0, 1, 0, D, N, 0, D * N };
            if (indP6 == 1)
            {
                for (int i = 0; i < 8; i++)
                {
                    TGk = TGk + X06[i] * DSTR13[i];
                }
            }
            else
            {
                for (int i = 0; i < 8; i++)
                {
                    TGk = TGk + X06[i] * DSTR14[i];
                }
            }
            return TGk;
        }
        public double dDk6dDp6(double Qt, double N)
        {
            double TGk = 0;
            double[] X06 = { 0, 1, 0, 0, Qt, 0, N, Qt * N };
            if (indP6 == 1)
            {
                for (int i = 0; i < 8; i++)
                {
                    TGk = TGk + X06[i] * DSTR13[i];
                }
            }
            else
            {
                for (int i = 0; i < 8; i++)
                {
                    TGk = TGk + X06[i] * DSTR14[i];
                }
            }
            return TGk;
        }
        public double dDk6dDp6dN6(double Qt)
        {
            double TGk = 0;
            double[] X06 = { 0, 0, 0, 0, 0, 0, 1, Qt };
            if (indP6 == 1)
            {
                for (int i = 0; i < 8; i++)
                {
                    TGk = TGk + X06[i] * DSTR13[i];
                }
            }
            else
            {
                for (int i = 0; i < 8; i++)
                {
                    TGk = TGk + X06[i] * DSTR14[i];
                }
            }
            return TGk;
        }
        public double dDk6dDp6dQt6(double N)
        {
            double TGk = 0;
            double[] X06 = { 0, 0, 0, 0, 1, 0, 0, N };
            if (indP6 == 1)
            {
                for (int i = 0; i < 8; i++)
                {
                    TGk = TGk + X06[i] * DSTR13[i];
                }
            }
            else
            {
                for (int i = 0; i < 8; i++)
                {
                    TGk = TGk + X06[i] * DSTR14[i];
                }
            }
            return TGk;
        }
        public double dDk6dQt6dN6(double D)
        {
            double TGk = 0;
            double[] X06 = { 0, 0, 0, 0, 0, 1, 0, D };
            if (indP6 == 1)
            {
                for (int i = 0; i < 8; i++)
                {
                    TGk = TGk + X06[i] * DSTR13[i];
                }
            }
            else
            {
                for (int i = 0; i < 8; i++)
                {
                    TGk = TGk + X06[i] * DSTR14[i];
                }
            }
            return TGk;
        }

        public double dGpvdN3()
        {
            double G = dQ3dN3(P3, Qt3, N3) / 0.595 / 0.995;
            return G;
        }
        public double dGpvdQt3()
        {
            double G = dQ3dQt3(P3, Qt3, N3) / 0.595 / 0.995;
            return G;
        }

        public double dNsntodN3()
        {
            double sn = 0;
            sn = -0.0000311462 * 2 * N3 + 0.0044165282;
            return sn;
        }

        public double dNsntodN3dN3()
        {
            double sn = 0;
            sn = -0.0000311462 * 2 + 0.0044165282;
            return sn;
        }
        public double dNsnpndQt3dQt3()
        {
            double npn = 0.0085088 * dQ3dQt3dQt3(P3, N3) / 0.595 / 0.995;
            return npn;
        }
        public double dNsnpndQt3dN3()
        {
            double npn = 0.0085088 * dQ3dQt3dN3(P3, Qt3, N3) / 0.595 / 0.995;
            return npn;
        }
        public double dNsnpndN3dN3()
        {
            double npn = 0.0085088 * (dQ3dNdN3(P3, Qt3) / 0.595 / 0.995);
            return npn;
        }
        public double dGpvdQt3dQt3()
        {
            double G = dQ3dQt3dQt3(P3, N3) / 0.595 / 0.995;
            return G;
        }
        public double dGpvdQt3dN3()
        {
            double G = dQ3dQt3dN3(P3, Qt3, N3) / 0.595 / 0.995;
            return G;
        }
        public double dGpvdN3dN3()
        {
            double G = dQ3dNdN3(P3, Qt3) / 0.595 / 0.995;
            return G;
        }

        public double dFdQpvk()
        {
            double F = -Ct * (dQptvm() / (7 * 0.98)) - Cnsn * dNsntdWB(Qptvm(Qpvk)) * dQptvm();
            return F;
        }
        public double dFdQpvkdQpvk()
        {
            double F = -Cnsn * ddNsntdWB() * Math.Pow(dQptvm(), 2);
            return F;
        }
        public double Qptvm(double Q)
        {
            double N = 0;
            if (Q != 0)
            {
                double[] Xp = { 1, Q };
                for (int i = 0; i < 2; i++)
                {
                    N = N + Xp[i] * DSTR47[i];
                }
            }
            return N;
        }
        public double dQptvm()
        {
            double N = 0;
            double[] Xp = { 0, 1 };
            for (int i = 0; i < 2; i++)
            {
                N = N + Xp[i] * DSTR47[i];
            }
            return N;
        }
        public double NsntdWB(double Q)
        {
            double N = 0;
            if (Q != 0)
            {
                double[] Xn = { 1, Q, Math.Pow(Q, 2) };
                for (int i = 0; i < 2; i++)
                {
                    N = N + Xn[i] * DSTR48[i];
                }
            }
            return N;
        }
        public double dNsntdWB(double Q)
        {
            double N = 0;
            double[] Xn = { 0, 1, 2 * Q };
            for (int i = 0; i < 2; i++)
            {
                N = N + Xn[i] * DSTR48[i];
            }
            return N;
        }
        public double ddNsntdWB()
        {
            double N = 0;
            double[] Xn = { 0, 0, 2 };
            for (int i = 0; i < 2; i++)
            {
                N = N + Xn[i] * DSTR48[i];
            }
            return N;
        }
        public double dLpdQpb()
        {
            double F = lp * 1 / kpb;
            return F;
        }

        public double dFdDptn()
        {
            double F = 0;
            F = -Cnsn * dNsnpndDptn();
            return F;
        }

        public double dLpdN1()
        {
            double F = lp * (dDauxdN1() + dD1dN1(P1, Qt1, N1) * (hv - h1) / (hv - h13));
            return F;
        }
        public double dLpdN1dN1()
        {
            double F = lp * (dDauxdN1dN1() + dD1dN1dN1(P1, Qt1) * (hv - h1) / (hv - h13));
            return F;
        }
        public double dLpdN1dQt1()
        {
            double F = lp * (dDauxdN1dQt1() + dD1dN1dQt1(P1, Qt1, N1) * (hv - h1) / (hv - h13));
            return F;
        }
        public double dLpdQt1()
        {
            double F = lp * (dDauxdQt1() + dD1dQt1(P1, Qt1, N1) * (hv - h1) / (hv - h13));
            return F;
        }
        public double dLpdQt1dQt1()
        {
            double F = lp * (dDauxdQt1dQt1() + dD1dQt1dQt1(P1, N1) * (hv - h1) / (hv - h13));
            return F;
        }

        public double dFdDp5()
        {
            double F = Cn * dNN5dDp() - Ct * (dQ5dD5() / (7 * 0.98 * nk / 100)) - Cnsn * (dNsnpndDp5() + dNsnko() * dGpvdDp5());
            return F;
        }
        public double dLp30dDp5()
        {
            double F = lp30 * (dDh5dDp5() - hv) / (hdop - hv);
            return F;
        }
        public double dNN5dDp()
        {
            double N = 0;
            double[] X5 = { 0, 1 };
            for (int i = 0; i < 2; i++)
            {
                N = N + X5[i] * DSTR46[i];
            }
            return N;
        }
        public double Q5(double D)
        {
            double Q = 0;
            if (indtg5 == 1)
            {
                if (D < 105) D = 105;
                double[] X5 = { 1, D };
                for (int i = 0; i < 2; i++)
                {
                    Q = Q + X5[i] * DSTR53[i];
                }
            }
            return Q;
        }
        public double dQ5dD5()
        {
            double Q = 0;
            double[] X5 = { 0, 1 };
            for (int i = 0; i < 2; i++)
            {
                Q = Q + X5[i] * DSTR53[i];
            }
            return Q;
        }

        public double Pk1(double Gk, double t)
        {
            double P = 0;
            if (Gk != 0)
            {
                double[] X01 = { 1, Gk, t, t * Gk, Math.Pow(Gk, 2), Math.Pow(t, 2), Math.Pow(Gk, 2) * t, Gk * Math.Pow(t, 2), Math.Pow(Gk, 2) * Math.Pow(t, 2) };
                if (Gcw1 == 3500)
                    for (int i = 0; i < 9; i++)
                    {
                        P = P + X01[i] * DSTR30[i];
                    }
                if (Gcw1 == 7000)
                    for (int i = 0; i < 9; i++)
                    {
                        P = P + X01[i] * DSTR31[i];
                    }
            }
            return P;
        }
        public double dPk1(double Gk, double t)
        {
            double P = 0;
            double[] X01 = { 0, 1, 0, t, 2 * Gk, 0, 2 * Gk * t, Math.Pow(t, 2), 2 * Gk * Math.Pow(t, 2) };
            if (Gcw1 == 3500)
                for (int i = 0; i < 9; i++)
                {
                    P = P + X01[i] * DSTR30[i];
                }
            if (Gcw1 == 7000)
                for (int i = 0; i < 9; i++)
                {
                    P = P + X01[i] * DSTR31[i];
                }
            return P;
        }
        public double ddPk1(double t)
        {
            double P = 0;
            double[] X01 = { 0, 0, 0, 0, 2, 0, 2 * t, 0, 2 * Math.Pow(t, 2) };
            if (Gcw1 == 3500)
                for (int i = 0; i < 9; i++)
                {
                    P = P + X01[i] * DSTR30[i];
                }
            if (Gcw1 == 7000)
                for (int i = 0; i < 9; i++)
                {
                    P = P + X01[i] * DSTR31[i];
                }
            return P;
        }

        public double dGpvdDp5()
        {
            double G = dQ5dD5() / 0.595 / 0.995;
            return G;
        }
        public double dLp30dDptn()
        {
            double F = -lp30 * (hdop - dDhptn()) / (hdop - h1);
            return F;
        }
        public double Dh5(double D)
        {
            double h = 0;
            if (D != 0)
            {
                h = 732.99 * D + 9300;
            }
            return h;
        }
        public double dDh5dDp5()
        {
            double h;
            h = 732.99;
            return h;
        }
        public double Dhptn(double D)
        {
            double h;
            if (D != 0)
            {
                h = 675.31 * D + 2468.1;
            }
            else h = 0;
            return h;
        }
        public double dDhptn()
        {
            double h = 675.31;
            return h;
        }

        public double h5(double D)
        {
            double h = 0.000383483 * Math.Pow(D, 2) - 0.357975195 * D + 837.4272844;
            return h;
        }

        public double D1(double P, double Qt, double N)
        {
            double D = 0;
            if (indtg1 == 1)
            {
                double[] X1 = { 1, P, Qt, N, P * Qt, Qt * N, P * N, P * Qt * N };
                for (int i = 0; i < 8; i++)
                {
                    D = D + X1[i] * DSTR1[i];
                }
            }
            return D;
        }
        public double dD1dN1(double P, double Qt, double N)
        {
            double D = 0;
            double[] X1 = { 0, 0, 0, 1, 0, Qt, P, P * Qt };
            for (int i = 0; i < 8; i++)
            {
                D = D + X1[i] * DSTR1[i];
            }
            return D;
        }
        public double dD1dN1dN1(double P, double Qt)
        {
            double D = 0;            
            return D;
        }
        public double dD1dN1dQt1(double P, double Qt, double N)
        {
            double D = 0;
            double[] X1 = { 0, 0, 0, 0, 0, 1, 0, P };
            for (int i = 0; i < 8; i++)
            {
                D = D + X1[i] * DSTR1[i];
            }
            return D;
        }
        public double dD1dQt1(double P, double Qt, double N)
        {
            double D = 0;
            double[] X1 = { 0, 0, 1, 0, P, N, 0, P * N };
            for (int i = 0; i < 8; i++)
            {
                D = D + X1[i] * DSTR1[i];
            }
            return D;
        }
        public double dD1dQt1dQt1(double P, double N)
        {
            double D = 0;            
            return D;
        }

        public double PoprTG1t(double t, double D)
        {
            double dN = 0;
            double[] X1 = { 1, t, D, t * D, Math.Pow(t, 2), Math.Pow(D, 2), Math.Pow(t, 2) * D, t * Math.Pow(D, 2), Math.Pow(t, 2) * Math.Pow(D, 2) };
            for (int i = 0; i < 9; i++)
            {
                dN = dN + X1[i] * DSTR56[i];
            }
            return dN;
        }

        public double D112() // 1.2 ата на СН от ТГ-1
        {
            double D1, G1, G3, G4, G6;
            G1 = 0; G3 = 0; G4 = 0; G6 = 0;
            if (indtg1 == 1) G1 = Dk1(P1, Qt1, N1);
            if (indtg3 == 1) G3 = Dk3(P3, Qt3, N3);
            if (indtg4 == 1) G4 = Dk4(P4, Qt4, N4);
            if (indtg6 == 1) G6 = Dk6(Dp6, Qt6, N6);
            double PV = ((Q2(Dp2) + Q3(P3, Qt3, N3) + Q4(P4, Qt4, N4) + Q5(Dp5) + Q6(Dp6, Qt6, N6)) / 0.595 + D140_13 + D140_30 + D140_45) / 0.995;
            Ddvd = 16 * PV / 535;
            Ddnd = Ghob * 64 / 536 - Ddvd * 0.01 - Q2(Dp2) * 0.01 - PV * 0.005 * 0.98;
            Dpsv = (tsw - tcwout(G1, G3, G4, G6, Gcw1, Gcw3, Gcw4, Gcw6, tnv)) * Gsw / 536;
            if (Dpsv < 0) Dpsv = 0;
            D1 = Ddnd + Dpsv;
            return D1;
        }

        public double dLpdDptn()
        {
            double F = -lp * (1 + (dDhptn() - h1) / (h1 - hdop)) * (hv - h1) / (hv - h13);
            return F;
        }
        public double dNsntodN1()//N1
        {
            double sn = -0.0000428233 * 2 * N1 + 0.0052171908;
            return sn;
        }
        public double dNsntodN1dN1()//N1
        {
            double sn = -0.0000428233 * 2;
            return sn;
        }

        public double Ogr3()
        {
            double F = 0;
            if (ogrN3 == 1) F = lN31 * (N3 - ogr_TG3r(P3, Qt3));
            if (ogrN3 == 2) F = lN31 * (N3 - ogr_TG3r(P3, Qt3)) + lN32 * (N3 - ogr_TG3l(P3, Qt3));
            if (ogrN3 == 3) F = lN32 * (N3 - ogr_TG3l(P3, Qt3));
            return F;
        }
        public double dOgrdN3()
        {
            double F = 0;
            if (ogrN3 == 1) F = lN31 * 1;
            if (ogrN3 == 2) F = lN31 * 1 + lN32 * 1;
            if (ogrN3 == 3) F = lN32 * 1;
            return F;
        }
        public double dOgrdN3dlN31()
        {
            double F = 0;
            if (ogrN3 == 1) F = 1;
            if (ogrN3 == 2) F = 1;
            if (ogrN3 == 3) F = 0;
            return F;
        }
        public double dOgrdN3dlN32()
        {
            double F = 0;
            if (ogrN3 == 1) F = 0;
            if (ogrN3 == 2) F = 1;
            if (ogrN3 == 3) F = 1;
            return F;
        }
        public double dOgrdQt3()
        {
            double F = 0;
            if (ogrN3 == 1) F = lN31 * -dogr_TG3r(P3, Qt3);
            if (ogrN3 == 2) F = lN31 * -dogr_TG3r(P3, Qt3) + lN32 * -dogr_TG3l(P3, Qt3);
            if (ogrN3 == 3) F = lN32 * -dogr_TG3l(P3, Qt3);
            return F;
        }
        public double dOgrdQt3dlN31()
        {
            double F = 0;
            if (ogrN3 == 1) F = -dogr_TG3r(P3, Qt3);
            if (ogrN3 == 2) F = -dogr_TG3r(P3, Qt3);
            if (ogrN3 == 3) F = 0;
            return F;
        }
        public double dOgrdQt3dlN32()
        {
            double F = 0;
            if (ogrN3 == 1) F = 0;
            if (ogrN3 == 2) F = -dogr_TG3l(P3, Qt3);
            if (ogrN3 == 3) F = -dogr_TG3l(P3, Qt3);
            return F;
        }
        public double dOgrdQt3dQt3()
        {
            double F = 0;
            if (ogrN3 == 1) F = lN31 * -ddogr_TG3r(P3, Qt3);
            if (ogrN3 == 2) F = lN31 * -ddogr_TG3r(P3, Qt3) + lN32 * -ddogr_TG3l(P3, Qt3);
            if (ogrN3 == 3) F = lN32 * -ddogr_TG3l(P3, Qt3);
            return F;
        }
        public double dOgr3dlN31()
        {
            double F = 0;
            if (ogrN3 == 1) F = N3 - ogr_TG3r(P3, Qt3);
            if (ogrN3 == 2) F = N3 - ogr_TG3r(P3, Qt3);
            if (ogrN3 == 3) F = 0;
            return F;
        }
        public double dOgr3dlN32()
        {
            double F = 0;
            if (ogrN3 == 1) F = 0;
            if (ogrN3 == 2) F = N3 - ogr_TG3l(P3, Qt3);
            if (ogrN3 == 3) F = N3 - ogr_TG3l(P3, Qt3);
            return F;
        }

        public double ogr_TG3r(double P, double Qt)
        {
            double N = 0;
            double[] X3 = {  1, P, Qt, P * Qt, Math.Pow(P, 2), Math.Pow(Qt, 2),
            Math.Pow(P, 2)*Qt, P*Math.Pow(Qt, 2),Math.Pow(P, 2)* Math.Pow(Qt, 2)};
            if (indP3 == 1)
            {
                for (int i = 0; i < 9; i++)
                {
                    N = N + X3[i] * DSTR18[i];
                }
            }
            else
            {
                for (int i = 0; i < 9; i++)
                {
                    N = N + X3[i] * DSTR19[i];
                }
            }
            return N;
        }
        public double dogr_TG3r(double P, double Qt)
        {
            double N = 0;
            double[] X3 = { 0, 0, 1, P, 0, 2 * Qt, Math.Pow(P, 2), P * 2 * Qt, Math.Pow(P, 2) * 2 * Qt };
            if (indP3 == 1)
            {
                for (int i = 0; i < 9; i++)
                {
                    N = N + X3[i] * DSTR18[i];
                }
            }
            else
            {
                for (int i = 0; i < 9; i++)
                {
                    N = N + X3[i] * DSTR19[i];
                }
            }
            return N;
        }
        public double ddogr_TG3r(double P, double Qt)
        {
            double N = 0;
            double[] X3 = { 0, 0, 0, 0, 0, 2, 0, P * 2, Math.Pow(P, 2) * 2 };
            if (indP3 == 1)
            {
                for (int i = 0; i < 9; i++)
                {
                    N = N + X3[i] * DSTR18[i];
                }
            }
            else
            {
                for (int i = 0; i < 9; i++)
                {
                    N = N + X3[i] * DSTR19[i];
                }
            }
            return N;
        }

        public double ogr_TG3l(double P, double Qt)
        {
            double N = 0;
            double[] X3 = {  1, P, Qt, P * Qt, Math.Pow(P, 2), Math.Pow(Qt, 2),
            Math.Pow(P, 2)*Qt, P*Math.Pow(Qt, 2),Math.Pow(P, 2)* Math.Pow(Qt, 2)};
            if (indP3 == 1)
            {
                for (int i = 0; i < 9; i++)
                {
                    N = N + X3[i] * DSTR20[i];
                }
            }
            else
            {
                for (int i = 0; i < 9; i++)
                {
                    N = N + X3[i] * DSTR21[i];
                }
            }
            return N;
        }
        public double dogr_TG3l(double P, double Qt)
        {
            double N = 0;
            double[] X3 = { 0, 0, 1, P, 0, 2 * Qt, Math.Pow(P, 2), P * 2 * Qt, Math.Pow(P, 2) * 2 * Qt };
            if (indP3 == 1)
            {
                for (int i = 0; i < 9; i++)
                {
                    N = N + X3[i] * DSTR20[i];
                }
            }
            else
            {
                for (int i = 0; i < 9; i++)
                {
                    N = N + X3[i] * DSTR21[i];
                }
            }
            return N;
        }
        public double ddogr_TG3l(double P, double Qt)
        {
            double N = 0;
            double[] X3 = { 0, 0, 0, 0, 0, 2, 0, P * 2, Math.Pow(P, 2) * 2 };
            if (indP3 == 1)
            {
                for (int i = 0; i < 9; i++)
                {
                    N = N + X3[i] * DSTR20[i];
                }
            }
            else
            {
                for (int i = 0; i < 9; i++)
                {
                    N = N + X3[i] * DSTR21[i];
                }
            }
            return N;
        }

        public double ogr_TG3Qtr(double P, double N)
        {
            double Q = 0;
            double[] X3 = {  1, P, N, P * N, Math.Pow(P, 2), Math.Pow(N, 2),
            Math.Pow(P, 2)*N, P*Math.Pow(N, 2),Math.Pow(P, 2)* Math.Pow(N, 2)};
            for (int i = 0; i < 9; i++)
            {
                Q = Q + X3[i] * DSTR66[i];
            }
            if (Q > Qt31)
            {
                Q = 0;
                for (int i = 0; i < 9; i++)
                {
                    Q = Q + X3[i] * DSTR65[i];
                }
            }
            return Q;
        }
        public double ogr_TG3Qtl(double P, double N)
        {
            double Q = 0;
            double[] X3 = {  1, P, N, P * N, Math.Pow(P, 2), Math.Pow(N, 2),
            Math.Pow(P, 2)*N, P*Math.Pow(N, 2),Math.Pow(P, 2)* Math.Pow(N, 2)};

            for (int i = 0; i < 9; i++)
            {
                Q = Q + X3[i] * DSTR64[i];
            }
            if (Q > Qt31)
            {
                Q = 0;
                for (int i = 0; i < 9; i++)
                {
                    Q = Q + X3[i] * DSTR63[i];
                }
            }
            return Q;
        }

        public double ogrTG1n(double P)
        {
            double N = -5.7127 * P + 31.255;
            return N;
        }
        public double ogrTG3n(double P, double Qt)
        {
            double N = 0;
            if (Qt > Qt31) N = -5 * P + 60;
            else N = -10 * P + 54;
            return N;
        }
        public double ogrTG4n(double P, double Qt)
        {
            double N = 0;
            if (Qt >= Qt41) N = -12.051 * P + 126.02;
            else N = -15 * P + 120.5;
            return N;
        }
        public double ogrTG6n(double D)
        {
            double N = 0;
            N = -0.0006813 * Math.Pow(D, 2) + 0.3602540 * D + 81.5705882;
            return N;
        }

        public double ogrTG6dn(double N)
        {
            double D = 5.4875 * N - 658.69;
            return D;
        } // огр. максимальный Dp от N справа, Dp<263
        public double ogrTG6db(double N)
        {
            double D = -4.9043 * N + 1087.9;
            return D;
        } // огр. максимальный Dp от N справа, Dp>263

        public double dOgrdN1()
        {
            double F = 0;
            if (ogrN1 == 1) F = lN11 * 1;
            if (ogrN1 == 2) F = lN11 * 1 + lN12 * 1;
            if (ogrN1 == 3) F = lN12 * 1;
            if (ogrN1 == 4) F = lN12 * 1 + lN13 * 1;
            if (ogrN1 == 5) F = lN13 * 1;
            return F;
        }
        public double dOgrdQt1()
        {
            double F = 0;
            if (ogrN1 == 1) F = lN11 * -dogr_TG1rb(P1, Qt1);
            if (ogrN1 == 2) F = lN11 * -dogr_TG1rb(P1, Qt1) + lN12 * -dogr_TG1rn(P1, Qt1);
            if (ogrN1 == 3) F = lN12 * -dogr_TG1rn(P1, Qt1);
            if (ogrN1 == 4) F = lN12 * -dogr_TG1rn(P1, Qt1) + lN13 * -dogr_TG1l(P1, Qt1);
            if (ogrN1 == 5) F = lN13 * -dogr_TG1l(P1, Qt1);
            return F;
        }
        public double dOgrdQt1dQt1()
        {
            double F = 0;
            if (ogrN1 == 1) F = lN11 * -ddogr_TG1rb(P1, Qt1);
            if (ogrN1 == 2) F = lN11 * -ddogr_TG1rb(P1, Qt1) + lN12 * -ddogr_TG1rn(P1, Qt1);
            if (ogrN1 == 3) F = lN12 * -ddogr_TG1rn(P1, Qt1);
            if (ogrN1 == 4) F = lN12 * -ddogr_TG1rn(P1, Qt1) + lN13 * -ddogr_TG1l(P1, Qt1);
            if (ogrN1 == 5) F = lN13 * -ddogr_TG1l(P1, Qt1);
            return F;
        }
        public double dOgr1dlN11()
        {
            double F = 0;
            if (ogrN1 == 1) F = N1 - ogr_TG1rb(P1, Qt1);
            if (ogrN1 == 2) F = N1 - ogr_TG1rb(P1, Qt1);
            if (ogrN1 == 3) F = 0;
            if (ogrN1 == 4) F = 0;
            if (ogrN1 == 5) F = 0;
            return F;
        }
        public double dOgr1dlN12()
        {
            double F = 0;
            if (ogrN1 == 1) F = 0;
            if (ogrN1 == 2) F = N1 - ogr_TG1rn(P1, Qt1);
            if (ogrN1 == 3) F = N1 - ogr_TG1rn(P1, Qt1);
            if (ogrN1 == 4) F = N1 - ogr_TG1rn(P1, Qt1);
            if (ogrN1 == 5) F = 0;
            return F;
        }
        public double dOgr1dlN13()
        {
            double F = 0;
            if (ogrN1 == 1) F = 0;
            if (ogrN1 == 2) F = 0;
            if (ogrN1 == 3) F = 0;
            if (ogrN1 == 4) F = N1 - ogr_TG1l(P1, Qt1);
            if (ogrN1 == 5) F = N1 - ogr_TG1l(P1, Qt1);
            return F;
        }
        public double dOgrdN1dlN11()
        {
            double F = 0;
            if (ogrN1 == 1) F = 1;
            if (ogrN1 == 2) F = 1;
            if (ogrN1 == 3) F = 0;
            if (ogrN1 == 4) F = 0;
            if (ogrN1 == 5) F = 0;
            return F;
        }
        public double dOgrdN1dlN12()
        {
            double F = 0;
            if (ogrN1 == 1) F = 0;
            if (ogrN1 == 2) F = 1;
            if (ogrN1 == 3) F = 1;
            if (ogrN1 == 4) F = 1;
            if (ogrN1 == 5) F = 0;
            return F;
        }
        public double dOgrdN1dlN13()
        {
            double F = 0;
            if (ogrN1 == 1) F = 0;
            if (ogrN1 == 2) F = 0;
            if (ogrN1 == 3) F = 0;
            if (ogrN1 == 4) F = 1;
            if (ogrN1 == 5) F = 1;
            return F;
        }
        public double dOgrdQt1dlN11()
        {
            double F = 0;
            if (ogrN1 == 1) F = -dogr_TG1rb(P1, Qt1);
            if (ogrN1 == 2) F = -dogr_TG1rb(P1, Qt1);
            if (ogrN1 == 3) F = 0;
            if (ogrN1 == 4) F = 0;
            if (ogrN1 == 5) F = 0;
            return F;
        }
        public double dOgrdQt1dlN12()
        {
            double F = 0;
            if (ogrN1 == 1) F = 0;
            if (ogrN1 == 2) F = -dogr_TG1rn(P1, Qt1);
            if (ogrN1 == 3) F = -dogr_TG1rn(P1, Qt1);
            if (ogrN1 == 4) F = -dogr_TG1rn(P1, Qt1);
            if (ogrN1 == 5) F = 0;
            return F;
        }
        public double dOgrdQt1dlN13()
        {
            double F = 0;
            if (ogrN1 == 1) F = 0;
            if (ogrN1 == 2) F = 0;
            if (ogrN1 == 3) F = 0;
            if (ogrN1 == 4) F = -dogr_TG1l(P1, Qt1);
            if (ogrN1 == 5) F = -dogr_TG1l(P1, Qt1);
            return F;
        }

        public double ogr_TG1rn(double P, double Qt)
        {
            double N = 0;
            double[] X1 = { 1, P, Qt, P * Qt, Math.Pow(P, 2), Math.Pow(Qt, 2), Math.Pow(P, 2) * Qt, P * Math.Pow(Qt, 2), Math.Pow(P, 2) * Math.Pow(Qt, 2) };
            for (int i = 0; i < 9; i++)
            {
                N = N + X1[i] * DSTR17[i];
            }
            return N;
        }
        public double dogr_TG1rn(double P, double Qt)
        {
            double N = 0;
            double[] X1 = { 0, 0, 1, P, 0, 2 * Qt, Math.Pow(P, 2), P * 2 * Qt, Math.Pow(P, 2) * 2 * Qt };
            for (int i = 0; i < 9; i++)
            {
                N = N + X1[i] * DSTR17[i];
            }
            return N;
        }
        public double ddogr_TG1rn(double P, double Qt)
        {
            double N = 0;
            double[] X1 = { 0, 0, 0, 0, 0, 2, 0, P * 2, Math.Pow(P, 2) * 2 };
            for (int i = 0; i < 9; i++)
            {
                N = N + X1[i] * DSTR17[i];
            }
            return N;
        }
        public double ogr_TG1rb(double P, double Qt)
        {
            double N = 0;
            double[] X1 = { 1, P, Qt, P * Qt, Math.Pow(P, 2), Math.Pow(Qt, 2), Math.Pow(P, 2) * Qt, P * Math.Pow(Qt, 2), Math.Pow(P, 2) * Math.Pow(Qt, 2) };
            for (int i = 0; i < 9; i++)
            {
                N = N + X1[i] * DSTR16[i];
            }
            return N;
        }
        public double dogr_TG1rb(double P, double Qt)
        {
            double N = 0;
            double[] X1 = { 0, 0, 1, P, 0, 2 * Qt, Math.Pow(P, 2), P * 2 * Qt, Math.Pow(P, 2) * 2 * Qt };
            for (int i = 0; i < 9; i++)
            {
                N = N + X1[i] * DSTR16[i];
            }
            return N;
        }
        public double ddogr_TG1rb(double P, double Qt)
        {
            double N = 0;
            double[] X1 = { 0, 0, 0, 0, 0, 2, 0, P * 2, Math.Pow(P, 2) * 2 };
            for (int i = 0; i < 9; i++)
            {
                N = N + X1[i] * DSTR16[i];
            }
            return N;
        }
        public double ogr_TG1l(double P, double Qt)
        {
            double N = 0;
            double[] X1 = { 1, P, Qt, P * Qt, Math.Pow(P, 2), Math.Pow(Qt, 2), Math.Pow(P, 2) * Qt, P * Math.Pow(Qt, 2), Math.Pow(P, 2) * Math.Pow(Qt, 2) };
            for (int i = 0; i < 9; i++)
            {
                N = N + X1[i] * DSTR15[i];
            }
            return N;
        }
        public double dogr_TG1l(double P, double Qt)
        {
            double N = 0;
            double[] X1 = { 0, 0, 1, P, 0, 2 * Qt, Math.Pow(P, 2), P * 2 * Qt, Math.Pow(P, 2) * 2 * Qt };
            for (int i = 0; i < 9; i++)
            {
                N = N + X1[i] * DSTR15[i];
            }
            return N;
        }
        public double ddogr_TG1l(double P, double Qt)
        {
            double N = 0;
            double[] X1 = { 0, 0, 0, 0, 0, 2, 0, P * 2, Math.Pow(P, 2) * 2 };
            for (int i = 0; i < 9; i++)
            {
                N = N + X1[i] * DSTR15[i];
            }
            return N;
        }

        public double ogr_TG1Qtrn(double P, double N)
        {
            double Q = 0;
            double[] X1 = { 1, P, N, P * N, Math.Pow(P, 2), Math.Pow(N, 2), Math.Pow(P, 2) * N, P * Math.Pow(N, 2), Math.Pow(P, 2) * Math.Pow(N, 2) };
            for (int i = 0; i < 9; i++)
            {
                Q = Q + X1[i] * DSTR62[i];
            }
            return Q;
        }
        public double ogr_TG1Qtrb(double P, double N)
        {
            double D = 0;
            double a, b, c;

            a = DSTR61[5] + DSTR61[7] * P + DSTR61[8] * Math.Pow(P, 2);
            b = DSTR61[2] + DSTR61[3] * P + DSTR61[6] * Math.Pow(P, 2);
            c = DSTR61[0] + DSTR61[1] * P + DSTR61[4] * Math.Pow(P, 2) - N;
            
            double Disc = Math.Pow(b, 2) - 4 * a * c;
            double x1, x2;
            x1 = (-b + Math.Pow(Disc, 0.5)) / (2 * a);
            x2 = (-b - Math.Pow(Disc, 0.5)) / (2 * a);
            if (x1 < x2) D = x1;
            else D = x2;            
            return D;            
        }
        public double ogr_TG1Qtl(double P, double N)
        {
            double Q = 0;
            double[] X1 = { 1, P, N, P * N, Math.Pow(P, 2), Math.Pow(N, 2), Math.Pow(P, 2) * N, P * Math.Pow(N, 2), Math.Pow(P, 2) * Math.Pow(N, 2) };
            for (int i = 0; i < 9; i++)
            {
                Q = Q + X1[i] * DSTR60[i];
            }
            return Q;
        }

        public double Ogr4()
        {
            double F = 0;
            if (ogrN4 == 1) F = lN41 * (N4 - 120);
            if (ogrN4 == 2) F = lN41 * (N4 - 120) + lN42 * (N4 - ogr_TG4r(P4, Qt4));
            if (ogrN4 == 3) F = lN42 * (N4 - ogr_TG4r(P4, Qt4));
            if (ogrN4 == 4) F = lN42 * (N4 - ogr_TG4r(P4, Qt4)) + lN43 * (N4 - ogr_TG4l(P4, Qt4));
            if (ogrN4 == 5) F = lN43 * (N4 - ogr_TG4l(P4, Qt4));
            if (ogrN4 == 6) F = lN43 * (N4 - ogr_TG4l(P4, Qt4)) + lN44 * (N4 - 40);
            if (ogrN4 == 7) F = lN44 * (N4 - 40);
            return F;
        }
        public double dOgrdN4()
        {
            double F = 0;
            if (ogrN4 == 1) F = lN41;
            if (ogrN4 == 2) F = lN41 + lN42;
            if (ogrN4 == 3) F = lN42;
            if (ogrN4 == 4) F = lN42 + lN43;
            if (ogrN4 == 5) F = lN43;
            if (ogrN4 == 6) F = lN43 + lN44;
            if (ogrN4 == 7) F = lN44;
            return F;
        }
        public double dOgrdQt4()
        {
            double F = 0;
            if (ogrN4 == 1) F = 0;
            if (ogrN4 == 2) F = lN42 * -dogr_TG4r(P4, Qt4);
            if (ogrN4 == 3) F = lN42 * -dogr_TG4r(P4, Qt4);
            if (ogrN4 == 4) F = lN42 * -dogr_TG4r(P4, Qt4) + lN43 * -dogr_TG4l(P4, Qt4);
            if (ogrN4 == 5) F = lN43 * -dogr_TG4l(P4, Qt4);
            if (ogrN4 == 6) F = lN43 * -dogr_TG4l(P4, Qt4);
            if (ogrN4 == 7) F = 0;
            return F;
        }
        public double dOgrdQt4dQt4()
        {
            double F = 0;
            if (ogrN4 == 1) F = 0;
            if (ogrN4 == 2) F = lN42 * -ddogr_TG4r(P4, Qt4);
            if (ogrN4 == 3) F = lN42 * -ddogr_TG4r(P4, Qt4);
            if (ogrN4 == 4) F = lN42 * -ddogr_TG4r(P4, Qt4) + lN43 * -ddogr_TG4l(P4, Qt4);
            if (ogrN4 == 5) F = lN43 * -ddogr_TG4l(P4, Qt4);
            if (ogrN4 == 6) F = lN43 * -ddogr_TG4l(P4, Qt4);
            if (ogrN4 == 7) F = 0;
            return F;
        }
        public double dOgr4dlN41()
        {
            double F = 0;
            if (ogrN4 == 1) F = N4 - 120;
            if (ogrN4 == 2) F = N4 - 120;
            if (ogrN4 == 3) F = 0;
            if (ogrN4 == 4) F = 0;
            if (ogrN4 == 5) F = 0;
            if (ogrN4 == 6) F = 0;
            if (ogrN4 == 7) F = 0;
            return F;
        }
        public double dOgr4dlN42()
        {
            double F = 0;
            if (ogrN4 == 1) F = 0;
            if (ogrN4 == 2) F = N4 - ogr_TG4r(P4, Qt4);
            if (ogrN4 == 3) F = N4 - ogr_TG4r(P4, Qt4);
            if (ogrN4 == 4) F = N4 - ogr_TG4r(P4, Qt4);
            if (ogrN4 == 5) F = 0;
            if (ogrN4 == 6) F = 0;
            if (ogrN4 == 7) F = 0;
            return F;
        }
        public double dOgr4dlN43()
        {
            double F = 0;
            if (ogrN4 == 1) F = 0;
            if (ogrN4 == 2) F = 0;
            if (ogrN4 == 3) F = 0;
            if (ogrN4 == 4) F = N4 - ogr_TG4l(P4, Qt4);
            if (ogrN4 == 5) F = N4 - ogr_TG4l(P4, Qt4);
            if (ogrN4 == 6) F = N4 - ogr_TG4l(P4, Qt4);
            if (ogrN4 == 7) F = 0;
            return F;
        }
        public double dOgr4dlN44()
        {
            double F = 0;
            if (ogrN4 == 1) F = 0;
            if (ogrN4 == 2) F = 0;
            if (ogrN4 == 3) F = 0;
            if (ogrN4 == 4) F = 0;
            if (ogrN4 == 5) F = 0;
            if (ogrN4 == 6) F = N4 - 40;
            if (ogrN4 == 7) F = N4 - 40;
            return F;
        }
        public double dOgrdN4dlN41()
        {
            double F = 0;
            if (ogrN4 == 1) F = 1;
            if (ogrN4 == 2) F = 1;
            if (ogrN4 == 3) F = 0;
            if (ogrN4 == 4) F = 0;
            if (ogrN4 == 5) F = 0;
            if (ogrN4 == 6) F = 0;
            if (ogrN4 == 7) F = 0;
            return F;
        }
        public double dOgrdN4dlN42()
        {
            double F = 0;
            if (ogrN4 == 1) F = 0;
            if (ogrN4 == 2) F = 1;
            if (ogrN4 == 3) F = 1;
            if (ogrN4 == 4) F = 1;
            if (ogrN4 == 5) F = 0;
            if (ogrN4 == 6) F = 0;
            if (ogrN4 == 7) F = 0;
            return F;
        }
        public double dOgrdN4dlN43()
        {
            double F = 0;
            if (ogrN4 == 1) F = 0;
            if (ogrN4 == 2) F = 0;
            if (ogrN4 == 3) F = 0;
            if (ogrN4 == 4) F = 1;
            if (ogrN4 == 5) F = 1;
            if (ogrN4 == 6) F = 1;
            if (ogrN4 == 7) F = 0;
            return F;
        }
        public double dOgrdN4dlN44()
        {
            double F = 0;
            if (ogrN4 == 1) F = 0;
            if (ogrN4 == 2) F = 0;
            if (ogrN4 == 3) F = 0;
            if (ogrN4 == 4) F = 0;
            if (ogrN4 == 5) F = 0;
            if (ogrN4 == 6) F = 1;
            if (ogrN4 == 7) F = 1;
            return F;
        }
        public double dOgrdQt4dlN41()
        {
            double F = 0;
            if (ogrN4 == 1) F = 0;
            if (ogrN4 == 2) F = 0;
            if (ogrN4 == 3) F = 0;
            if (ogrN4 == 4) F = 0;
            if (ogrN4 == 5) F = 0;
            if (ogrN4 == 6) F = 0;
            if (ogrN4 == 7) F = 0;
            return F;
        }
        public double dOgrdQt4dlN42()
        {
            double F = 0;
            if (ogrN4 == 1) F = 0;
            if (ogrN4 == 2) F = -dogr_TG4r(P4, Qt4);
            if (ogrN4 == 3) F = -dogr_TG4r(P4, Qt4);
            if (ogrN4 == 4) F = -dogr_TG4r(P4, Qt4);
            if (ogrN4 == 5) F = 0;
            if (ogrN4 == 6) F = 0;
            if (ogrN4 == 7) F = 0;
            return F;
        }
        public double dOgrdQt4dlN43()
        {
            double F = 0;
            if (ogrN4 == 1) F = 0;
            if (ogrN4 == 2) F = 0;
            if (ogrN4 == 3) F = 0;
            if (ogrN4 == 4) F = -dogr_TG4l(P4, Qt4);
            if (ogrN4 == 5) F = -dogr_TG4l(P4, Qt4);
            if (ogrN4 == 6) F = -dogr_TG4l(P4, Qt4);
            if (ogrN4 == 7) F = 0;
            return F;
        }
        public double dOgrdQt4dlN44()
        {
            double F = 0;
            if (ogrN4 == 1) F = 0;
            if (ogrN4 == 2) F = 0;
            if (ogrN4 == 3) F = 0;
            if (ogrN4 == 4) F = 0;
            if (ogrN4 == 5) F = 0;
            if (ogrN4 == 6) F = 0;
            if (ogrN4 == 7) F = 0;
            return F;
        }

        public double ogr_TG4l(double P, double Qt)
        {
            double N = 0;
            double[] X4 = {  1, P, Qt, P * Qt, Math.Pow(P, 2), Math.Pow(Qt, 2),
            Math.Pow(P, 2)*Qt, P*Math.Pow(Qt, 2),Math.Pow(P, 2)* Math.Pow(Qt, 2)};
            if (indP4 == 1)
            {
                for (int i = 0; i < 9; i++)
                {
                    N = N + X4[i] * DSTR24[i];
                }
            }
            else
            {
                for (int i = 0; i < 9; i++)
                {
                    N = N + X4[i] * DSTR25[i];
                }
            }
            return N;
        }
        public double dogr_TG4l(double P, double Qt)
        {
            double N = 0;
            double[] X4 = { 0, 0, 1, P, 0, 2 * Qt, Math.Pow(P, 2), P * 2 * Qt, Math.Pow(P, 2) * 2 * Qt };
            if (indP4 == 1)
            {
                for (int i = 0; i < 9; i++)
                {
                    N = N + X4[i] * DSTR24[i];
                }
            }
            else
            {
                for (int i = 0; i < 9; i++)
                {
                    N = N + X4[i] * DSTR25[i];
                }
            }
            return N;
        }
        public double ddogr_TG4l(double P, double Qt)
        {
            double N = 0;
            double[] X4 = { 0, 0, 0, 0, 0, 2, 0, P * 2, Math.Pow(P, 2) * 2 };
            if (indP4 == 1)
            {
                for (int i = 0; i < 9; i++)
                {
                    N = N + X4[i] * DSTR24[i];
                }
            }
            else
            {
                for (int i = 0; i < 9; i++)
                {
                    N = N + X4[i] * DSTR25[i];
                }
            }
            return N;
        }
        public double ogr_TG4r(double P, double Qt)
        {
            double N = 0;
            double[] X4 = {  1, P, Qt, P * Qt, Math.Pow(P, 2), Math.Pow(Qt, 2),
            Math.Pow(P, 2)*Qt, P*Math.Pow(Qt, 2),Math.Pow(P, 2)* Math.Pow(Qt, 2)};
            if (indP4 == 1)
            {
                for (int i = 0; i < 9; i++)
                {
                    N = N + X4[i] * DSTR22[i];
                }
            }
            else
            {
                for (int i = 0; i < 9; i++)
                {
                    N = N + X4[i] * DSTR23[i];
                }
            }
            return N;
        }
        public double dogr_TG4r(double P, double Qt)
        {
            double N = 0;
            double[] X4 = { 0, 0, 1, P, 0, 2 * Qt, Math.Pow(P, 2), P * 2 * Qt, Math.Pow(P, 2) * 2 * Qt };
            if (indP4 == 1)
            {
                for (int i = 0; i < 9; i++)
                {
                    N = N + X4[i] * DSTR22[i];
                }
            }
            else
            {
                for (int i = 0; i < 9; i++)
                {
                    N = N + X4[i] * DSTR23[i];
                }
            }
            return N;
        }
        public double ddogr_TG4r(double P, double Qt)
        {
            double N = 0;
            double[] X4 = { 0, 0, 0, 0, 0, 2, 0, P * 2, Math.Pow(P, 2) * 2 };
            if (indP4 == 1)
            {
                for (int i = 0; i < 9; i++)
                {
                    N = N + X4[i] * DSTR22[i];
                }
            }
            else
            {
                for (int i = 0; i < 9; i++)
                {
                    N = N + X4[i] * DSTR23[i];
                }
            }
            return N;
        }

        public double ogr_TG4Qtr(double P, double N)
        {
            double D = 0;
            double a, b, c;
            if (indP4 == 1)
            {
                a = DSTR22[5] + DSTR22[7] * P + DSTR22[8] * Math.Pow(P, 2);
                b = DSTR22[2] + DSTR22[3] * P + DSTR22[6] * Math.Pow(P, 2);
                c = DSTR22[0] + DSTR22[1] * P + DSTR22[4] * Math.Pow(P, 2) - N;                                
            }
            else
            {
                a = DSTR23[5] + DSTR23[7] * P + DSTR23[8] * Math.Pow(P, 2);
                b = DSTR23[2] + DSTR23[3] * P + DSTR23[6] * Math.Pow(P, 2);
                c = DSTR23[0] + DSTR23[1] * P + DSTR23[4] * Math.Pow(P, 2) - N; 
            }
            double Disc = Math.Pow(b, 2) - 4 * a * c;
            double x1, x2;
            x1 = (-b + Math.Pow(Disc, 0.5)) / (2 * a);
            x2 = (-b - Math.Pow(Disc, 0.5)) / (2 * a);
            if (x1 < -10 || x1 > 170) D = x2;
            else
            {
                if (x2 > -10 || x2 < 170)
                {
                    if (x2 > x1) D = x2;
                    else D = x1;
                }
                D = x1;
            }
            return D;
        }

        public double ogr_TG4Qtl(double P, double N)
        {
            double D = 0;
            double a, b, c;
            if (indP4 == 1)
            {
                a = DSTR24[5] + DSTR24[7] * P + DSTR24[8] * Math.Pow(P, 2);
                b = DSTR24[2] + DSTR24[3] * P + DSTR24[6] * Math.Pow(P, 2);
                c = DSTR24[0] + DSTR24[1] * P + DSTR24[4] * Math.Pow(P, 2) - N;
            }
            else
            {
                a = DSTR25[5] + DSTR25[7] * P + DSTR25[8] * Math.Pow(P, 2);
                b = DSTR25[2] + DSTR25[3] * P + DSTR25[6] * Math.Pow(P, 2);
                c = DSTR25[0] + DSTR25[1] * P + DSTR25[4] * Math.Pow(P, 2) - N;
            }
            double Disc = Math.Pow(b, 2) - 4 * a * c;
            double x1, x2;
            x1 = (-b + Math.Pow(Disc, 0.5)) / (2 * a);
            x2 = (-b - Math.Pow(Disc, 0.5)) / (2 * a);
            if (x1 < -10 || x1 > 170) D = x2;
            else
            {
                if (x2 > -10 || x2 < 170)
                {
                    if (x2 > x1) D = x2;
                    else D = x1;
                }
                D = x1;
            }
            return D;
        }

        public double Ogr6()
        {
            double F = 0;
            if (ogrN6 == 1) F = lN61 * (N6 - ogr_TG6r(Dp6, Qt6));
            if (ogrN6 == 2) F = lN61 * (N6 - ogr_TG6r(Dp6, Qt6)) + lN62 * (N6 - ogr_TG6ln(Dp6, Qt6));
            if (ogrN6 == 3) F = lN62 * (N6 - ogr_TG6ln(Dp6, Qt6));
            if (ogrN6 == 4) F = lN62 * (N6 - ogr_TG6ln(Dp6, Qt6)) + lN63 * (N6 - ogr_TG6lb(Dp6, Qt6));
            if (ogrN6 == 5) F = lN63 * (N6 - ogr_TG6lb(Dp6, Qt6));
            return F;
        }
        public double dOgrdDp6()
        {
            double F = 0;
            if (ogrN6 == 1) F = lN61 * -dogr_TG6rdDp6(Dp6, Qt6);
            if (ogrN6 == 2) F = lN61 * -dogr_TG6rdDp6(Dp6, Qt6) + lN62 * -dogr_TG6lndDp6(Dp6, Qt6);
            if (ogrN6 == 3) F = lN62 * -dogr_TG6lndDp6(Dp6, Qt6);
            if (ogrN6 == 4) F = lN62 * -dogr_TG6lndDp6(Dp6, Qt6) + lN63 * -dogr_TG6lbdDp6(Dp6, Qt6);
            if (ogrN6 == 5) F = lN63 * -dogr_TG6lbdDp6(Dp6, Qt6);
            return F;
        }
        public double dOgrdDp6dDp6()
        {
            double F = 0;
            if (ogrN6 == 1) F = lN61 * -ddogr_TG6rdDp6(Dp6, Qt6);
            if (ogrN6 == 2) F = lN61 * -ddogr_TG6rdDp6(Dp6, Qt6) + lN62 * -ddogr_TG6lndDp6(Dp6, Qt6);
            if (ogrN6 == 3) F = lN62 * -ddogr_TG6lndDp6(Dp6, Qt6);
            if (ogrN6 == 4) F = lN62 * -ddogr_TG6lndDp6(Dp6, Qt6) + lN63 * -ddogr_TG6lbdDp6(Dp6, Qt6);
            if (ogrN6 == 5) F = lN63 * -ddogr_TG6lbdDp6(Dp6, Qt6);
            return F;
        }
        public double dOgrdDp6dQt6()
        {
            double F = 0;
            if (ogrN6 == 1) F = lN61 * -dogr_TG6rdDp6dQt6(Dp6, Qt6);
            if (ogrN6 == 2) F = lN61 * -dogr_TG6rdDp6dQt6(Dp6, Qt6) + lN62 * -dogr_TG6lndDp6dQt6(Dp6, Qt6);
            if (ogrN6 == 3) F = lN62 * -dogr_TG6lndDp6dQt6(Dp6, Qt6);
            if (ogrN6 == 4) F = lN62 * -dogr_TG6lndDp6dQt6(Dp6, Qt6) + lN63 * -dogr_TG6lbdDp6dQt6(Dp6, Qt6);
            if (ogrN6 == 5) F = lN63 * -dogr_TG6lbdDp6dQt6(Dp6, Qt6);
            return F;
        }
        public double dOgrdDp6dlN61()
        {
            double F = 0;
            if (ogrN6 == 1) F = -dogr_TG6rdDp6(Dp6, Qt6);
            if (ogrN6 == 2) F = -dogr_TG6rdDp6(Dp6, Qt6);
            if (ogrN6 == 3) F = 0;
            if (ogrN6 == 4) F = 0;
            if (ogrN6 == 5) F = 0;
            return F;
        }
        public double dOgrdDp6dlN62()
        {
            double F = 0;
            if (ogrN6 == 1) F = 0;
            if (ogrN6 == 2) F = -dogr_TG6lndDp6(Dp6, Qt6);
            if (ogrN6 == 3) F = -dogr_TG6lndDp6(Dp6, Qt6);
            if (ogrN6 == 4) F = -dogr_TG6lndDp6(Dp6, Qt6);
            if (ogrN6 == 5) F = 0;
            return F;
        }
        public double dOgrdDp6dlN63()
        {
            double F = 0;
            if (ogrN6 == 1) F = 0;
            if (ogrN6 == 2) F = 0;
            if (ogrN6 == 3) F = 0;
            if (ogrN6 == 4) F = -dogr_TG6lbdDp6(Dp6, Qt6);
            if (ogrN6 == 5) F = -dogr_TG6lbdDp6(Dp6, Qt6);
            return F;
        }

        public double dOgrdN6()
        {
            double F = 0;
            if (ogrN6 == 1) F = lN61 * 1;
            if (ogrN6 == 2) F = lN61 * 1 + lN62 * 1;
            if (ogrN6 == 3) F = lN62 * 1;
            if (ogrN6 == 4) F = lN62 * 1 + lN63 * 1;
            if (ogrN6 == 5) F = lN63 * 1;
            return F;
        }
        public double dOgrdQt6()
        {
            double F = 0;
            if (ogrN6 == 1) F = lN61 * -dogr_TG6rdQt6(Dp6, Qt6);
            if (ogrN6 == 2) F = lN61 * -dogr_TG6rdQt6(Dp6, Qt6) + lN62 * -dogr_TG6lndQt6(Dp6, Qt6);
            if (ogrN6 == 3) F = lN62 * -dogr_TG6lndQt6(Dp6, Qt6);
            if (ogrN6 == 4) F = lN62 * -dogr_TG6lndQt6(Dp6, Qt6) + lN63 * -dogr_TG6lbdQt6(Dp6, Qt6);
            if (ogrN6 == 5) F = lN63 * -dogr_TG6lbdQt6(Dp6, Qt6);
            return F;
        }
        public double dOgrdQt6dQt6()
        {
            double F = 0;
            if (ogrN6 == 1) F = lN61 * -ddogr_TG6rdQt6(Dp6, Qt6);
            if (ogrN6 == 2) F = lN61 * -ddogr_TG6rdQt6(Dp6, Qt6) + lN62 * -ddogr_TG6lndQt6(Dp6, Qt6);
            if (ogrN6 == 3) F = lN62 * -ddogr_TG6lndQt6(Dp6, Qt6);
            if (ogrN6 == 4) F = lN62 * -ddogr_TG6lndQt6(Dp6, Qt6) + lN63 * -ddogr_TG6lbdQt6(Dp6, Qt6);
            if (ogrN6 == 5) F = lN63 * -ddogr_TG6lbdQt6(Dp6, Qt6);
            return F;
        }
        public double dOgr6dlN61()
        {
            double F = 0;
            if (ogrN6 == 1) F = N6 - ogr_TG6r(Dp6, Qt6);
            if (ogrN6 == 2) F = N6 - ogr_TG6r(Dp6, Qt6);
            if (ogrN6 == 3) F = 0;
            if (ogrN6 == 4) F = 0;
            if (ogrN6 == 5) F = 0;
            return F;
        }
        public double dOgr6dlN62()
        {
            double F = 0;
            if (ogrN6 == 1) F = 0;
            if (ogrN6 == 2) F = N6 - ogr_TG6ln(Dp6, Qt6);
            if (ogrN6 == 3) F = N6 - ogr_TG6ln(Dp6, Qt6);
            if (ogrN6 == 4) F = N6 - ogr_TG6ln(Dp6, Qt6);
            if (ogrN6 == 5) F = 0;
            return F;
        }
        public double dOgr6dlN63()
        {
            double F = 0;
            if (ogrN6 == 1) F = 0;
            if (ogrN6 == 2) F = 0;
            if (ogrN6 == 3) F = 0;
            if (ogrN6 == 4) F = N6 - ogr_TG6lb(Dp6, Qt6);
            if (ogrN6 == 5) F = N6 - ogr_TG6lb(Dp6, Qt6);
            return F;
        }
        public double dOgrdN6dlN61()
        {
            double F = 0;
            if (ogrN6 == 1) F = 1;
            if (ogrN6 == 2) F = 1;
            if (ogrN6 == 3) F = 0;
            if (ogrN6 == 4) F = 0;
            if (ogrN6 == 5) F = 0;
            return F;
        }
        public double dOgrdN6dlN62()
        {
            double F = 0;
            if (ogrN6 == 1) F = 0;
            if (ogrN6 == 2) F = 1;
            if (ogrN6 == 3) F = 1;
            if (ogrN6 == 4) F = 1;
            if (ogrN6 == 5) F = 0;
            return F;
        }
        public double dOgrdN6dlN63()
        {
            double F = 0;
            if (ogrN6 == 1) F = 0;
            if (ogrN6 == 2) F = 0;
            if (ogrN6 == 3) F = 0;
            if (ogrN6 == 4) F = 1;
            if (ogrN6 == 5) F = 1;
            return F;
        }
        public double dOgrdQt6dlN61()
        {
            double F = 0;
            if (ogrN6 == 1) F = -dogr_TG6r(Dp6, Qt6);
            if (ogrN6 == 2) F = -dogr_TG6r(Dp6, Qt6);
            if (ogrN6 == 3) F = 0;
            if (ogrN6 == 4) F = 0;
            if (ogrN6 == 5) F = 0;
            return F;
        }
        public double dOgrdQt6dlN62()
        {
            double F = 0;
            if (ogrN6 == 1) F = 0;
            if (ogrN6 == 2) F = -dogr_TG6lndQt6(Dp6, Qt6);
            if (ogrN6 == 3) F = -dogr_TG6lndQt6(Dp6, Qt6);
            if (ogrN6 == 4) F = -dogr_TG6lndQt6(Dp6, Qt6);
            if (ogrN6 == 5) F = 0;
            return F;
        }
        public double dOgrdQt6dlN63()
        {
            double F = 0;
            if (ogrN6 == 1) F = 0;
            if (ogrN6 == 2) F = 0;
            if (ogrN6 == 3) F = 0;
            if (ogrN6 == 4) F = -dogr_TG6lbdQt6(Dp6, Qt6);
            if (ogrN6 == 5) F = -dogr_TG6lbdQt6(Dp6, Qt6);
            return F;
        }

        public double ogr_TG6bt(double Qt)
        {
            double Qp = 0;
            double[] X6 = { 1, Qt, Math.Pow(Qt, 2) };
            for (int i = 0; i < 3; i++)
            {
                Qp = Qp + X6[i] * DSTR42[i];
            }
            return Qp;
        } // огр. Dp от Qt снизу
        public double ogr_TG6Dp(double Dp)
        {
            double Qt = 0;
            double[] X6 = { 1, Dp, Math.Pow(Dp, 2), Math.Pow(Dp, 3), Math.Pow(Dp, 4), Math.Pow(Dp, 5), Math.Pow(Dp, 6) };
            for (int i = 0; i < 7; i++)
            {
                Qt = Qt + X6[i] * DSTR43[i];
            }
            return Qt;
        } // огр. Qt от Dp слева  
        public double ogr_TG6Qt(double Qt)
        {
            double Dp = 0;
            Dp = -0.02153 * Math.Pow(Qt6, 3) + 3.3148 * Math.Pow(Qt6, 2) - 167.48993 * Qt6 + 2790.65193;
            return Dp;
        } // огр. Dp от Qt слева  
        public double ogr_TG6N(double Dp)
        {
            double N = 0;
            double[] X6 = { 1, Dp, Math.Pow(Dp, 2) };
            for (int i = 0; i < 2; i++)
            {
                N = N + X6[i] * DSTR71[i];
            }
            return N;
        }   // огр. минимальный N от Dp слева  
        public double ogr_TG6D(double N)
        {
            double D = 0;
            D = 5.7982 * N - 88.463;
            return D;
        }   // огр. минимальный Dp от N слева  
        public double ogr_TG6bd(double Dp)
        {
            double Qt = 0;
            Qt = -0.44701 * Dp + 268.31771;
            return Qt;
        } // огр. Qt от Dp снизу

        public double ogr_TG1Qt(double P)
        {
            double Qt = 0;
            double[] X = { 1, P, Math.Pow(P, 2) };
            for (int i = 0; i < 3; i++)
            {
                Qt = Qt + X[i] * DSTR44[i];
            }
            return Qt;
        }

        public double Kogr_TG6l(double Dp, double Qt)
        {
            double N = 0;
            if (Qt < ogr_TG6Dp(Dp)) { N = ogr_TG6lb(Dp, Qt); }
            if (Qt >= ogr_TG6Dp(Dp)) { N = ogr_TG6ln(Dp, Qt); }
            return N;
        }
        public double ogr_TG6ln(double Dp, double Qt)
        {
            double N = 0;
            double[] X6 = { 1, Dp, Qt, Dp * Qt };
            for (int i = 0; i < 4; i++)
            {
                N = N + X6[i] * DSTR27[i];
            }
            return N;
        }
        public double ogr_TG6lb(double Dp, double Qt)
        {
            double N = 0;
            double[] X6 = { 1, Dp, Qt, Dp * Qt, Math.Pow(Dp, 2), Math.Pow(Qt, 2),
            Math.Pow(Dp, 2)*Qt, Dp*Math.Pow(Qt, 2),Math.Pow(Dp, 2)* Math.Pow(Qt, 2)};
            for (int i = 0; i < 9; i++)
            {
                N = N + X6[i] * DSTR26[i];
            }
            return N;
        }
        public double ogr_TG6r(double Qp, double Qt)
        {
            double N = 0;
            double[] X6 = { 1, Qp, Qt, Qp * Qt, Math.Pow(Qp, 2), Math.Pow(Qt, 2),
            Math.Pow(Qp, 2)*Qt, Qp*Math.Pow(Qt, 2),Math.Pow(Qp, 2)* Math.Pow(Qt, 2)};
            for (int i = 0; i < 9; i++)
            {
                N = N + X6[i] * DSTR28[i];
            }
            return N;
        }

        public double dogr_TG6rdDp6(double Qp, double Qt)
        {
            double N = 0;
            double[] X6 = { 0, 1, 0, Qt, 2 * Qp, 0, 2 * Qp * Qt, Math.Pow(Qt, 2), 2 * Qp * Math.Pow(Qt, 2) };
            for (int i = 0; i < 9; i++)
            {
                N = N + X6[i] * DSTR28[i];
            }
            return N;
        }
        public double ddogr_TG6rdDp6(double Qp, double Qt)
        {
            double N = 0;
            double[] X6 = { 0, 0, 0, 0, 2, 0, 2 * Qt, 0, 2 * Math.Pow(Qt, 2) };
            for (int i = 0; i < 9; i++)
            {
                N = N + X6[i] * DSTR28[i];
            }
            return N;
        }

        public double dogr_TG6rdDp6dQt6(double Qp, double Qt)
        {
            double N = 0;
            double[] X6 = { 0, 0, 0, 1, 0, 0, 2 * Qp, 2 * Qt, 2 * Qp * 2 * Qt };
            for (int i = 0; i < 9; i++)
            {
                N = N + X6[i] * DSTR28[i];
            }
            return N;
        }
        public double dogr_TG6lndDp6dQt6(double Qp, double Qt)
        {
            double N = 0;
            double[] X6 = { 0, 0, 0, 1, 0, 0, 2 * Qp, 2 * Qt, 2 * Qp * 2 * Qt };
            for (int i = 0; i < 9; i++)
            {
                N = N + X6[i] * DSTR27[i];
            }
            return N;
        }
        public double dogr_TG6lbdDp6dQt6(double Qp, double Qt)
        {
            double N = 0;
            double[] X6 = { 0, 0, 0, 1, 0, 0, 2 * Qp, 2 * Qt, 2 * Qp * 2 * Qt };
            for (int i = 0; i < 9; i++)
            {
                N = N + X6[i] * DSTR26[i];
            }
            return N;
        }

        public double dogr_TG6lndDp6(double Qp, double Qt)
        {
            double N = 0;
            double[] X6 = { 0, 1, 0, Qt, 2 * Qp, 0, 2 * Qp * Qt, Math.Pow(Qt, 2), 2 * Qp * Math.Pow(Qt, 2) };
            for (int i = 0; i < 9; i++)
            {
                N = N + X6[i] * DSTR27[i];
            }
            return N;
        }
        public double ddogr_TG6lndDp6(double Qp, double Qt)
        {
            double N = 0;
            double[] X6 = { 0, 0, 0, 0, 2, 0, 2 * Qt, 0, 2 * Math.Pow(Qt, 2) };
            for (int i = 0; i < 9; i++)
            {
                N = N + X6[i] * DSTR27[i];
            }
            return N;
        }
        public double dogr_TG6lbdDp6(double Qp, double Qt)
        {
            double N = 0;
            double[] X6 = { 0, 1, 0, Qt, 2 * Qp, 0, 2 * Qp * Qt, Math.Pow(Qt, 2), 2 * Qp * Math.Pow(Qt, 2) };
            for (int i = 0; i < 9; i++)
            {
                N = N + X6[i] * DSTR26[i];
            }
            return N;
        }
        public double ddogr_TG6lbdDp6(double Qp, double Qt)
        {
            double N = 0;
            double[] X6 = { 0, 0, 0, 0, 2, 0, 2 * Qt, 0, 2 * Math.Pow(Qt, 2) };
            for (int i = 0; i < 9; i++)
            {
                N = N + X6[i] * DSTR26[i];
            }
            return N;
        }
        public double dogr_TG6r(double Qp, double Qt)
        {
            double N = 0;
            double[] X6 = { 0, 0, 1, Qp, 0, 2 * Qt, Math.Pow(Qp, 2), Qp * 2 * Qt, Math.Pow(Qp, 2) * 2 * Qt };
            for (int i = 0; i < 9; i++)
            {
                N = N + X6[i] * DSTR28[i];
            }
            return N;
        }
        public double ddogr_TG6r(double Qp, double Qt)
        {
            double N = 0;
            double[] X6 = { 0, 0, 0, 0, 0, 2, 0, Qp * 2, Math.Pow(Qp, 2) * 2 };
            for (int i = 0; i < 9; i++)
            {
                N = N + X6[i] * DSTR28[i];
            }
            return N;
        }

        public double dogr_TG6rdQt6(double Qp, double Qt)
        {
            double N = 0;
            double[] X6 = { 0, 0, 1, Qp, 0, 2 * Qt, Math.Pow(Qp, 2), Qp * 2 * Qt, Math.Pow(Qp, 2) * 2 * Qt };
            for (int i = 0; i < 9; i++)
            {
                N = N + X6[i] * DSTR28[i];
            }
            return N;
        }
        public double ddogr_TG6rdQt6(double Qp, double Qt)
        {
            double N = 0;
            double[] X6 = { 0, 0, 0, 0, 0, 2, 0, Qp * 2, Math.Pow(Qp, 2) * 2 };
            for (int i = 0; i < 9; i++)
            {
                N = N + X6[i] * DSTR28[i];
            }
            return N;
        }
        public double dogr_TG6lndQt6(double Qp, double Qt)
        {
            double N = 0;
            double[] X6 = { 0, 0, 1, Qp, 0, 2 * Qt, Math.Pow(Qp, 2), Qp * 2 * Qt, Math.Pow(Qp, 2) * 2 * Qt };
            for (int i = 0; i < 9; i++)
            {
                N = N + X6[i] * DSTR27[i];
            }
            return N;
        }
        public double ddogr_TG6lndQt6(double Qp, double Qt)
        {
            double N = 0;
            double[] X6 = { 0, 0, 0, 0, 0, 2, 0, Qp * 2, Math.Pow(Qp, 2) * 2 };
            for (int i = 0; i < 9; i++)
            {
                N = N + X6[i] * DSTR27[i];
            }
            return N;
        }
        public double dogr_TG6lbdQt6(double Qp, double Qt)
        {
            double N = 0;
            double[] X6 = { 0, 0, 1, Qp, 0, 2 * Qt, Math.Pow(Qp, 2), Qp * 2 * Qt, Math.Pow(Qp, 2) * 2 * Qt };
            for (int i = 0; i < 9; i++)
            {
                N = N + X6[i] * DSTR26[i];
            }
            return N;
        }
        public double ddogr_TG6lbdQt6(double Qp, double Qt)
        {
            double N = 0;
            double[] X6 = { 0, 0, 0, 0, 0, 2, 0, Qp * 2, Math.Pow(Qp, 2) * 2 };
            for (int i = 0; i < 9; i++)
            {
                N = N + X6[i] * DSTR26[i];
            }
            return N;
        }

        public double ogr_TG6Qtln(double Dp, double N)
        {
            double Q = (N - DSTR27[0] - DSTR27[1] * Dp) / (DSTR27[2] + DSTR27[3] * Dp);
            return Q;
        }
        public double ogr_TG6Qtlb(double Dp, double N)
        {
            double D = 0;
            double a, b, c;

            a = DSTR26[5] + DSTR26[7] * Dp + DSTR26[8] * Math.Pow(Dp, 2);
            b = DSTR26[2] + DSTR26[3] * Dp + DSTR26[6] * Math.Pow(Dp, 2);
            c = DSTR26[0] + DSTR26[1] * Dp + DSTR26[4] * Math.Pow(Dp, 2) - N;

            double Disc = Math.Pow(b, 2) - 4 * a * c;
            double x1, x2;
            x1 = (-b + Math.Pow(Disc, 0.5)) / (2 * a);
            x2 = (-b - Math.Pow(Disc, 0.5)) / (2 * a);
            if (x1 > -10 & x1 < 170)
            {
                if (x2 > -10 & x2 < 170)
                {
                    if (x2 > x1) D = x2;
                    else D = x1;
                }
                else D = x1;
            }
            else D = x2;
            return D;
        }
        public double ogr_TG6Qtr(double Dp, double N)
        {
            double D = 0;
            double a, b, c;

            a = DSTR28[5] + DSTR28[7] * Dp + DSTR28[8] * Math.Pow(Dp, 2);
            b = DSTR28[2] + DSTR28[3] * Dp + DSTR28[6] * Math.Pow(Dp, 2);
            c = DSTR28[0] + DSTR28[1] * Dp + DSTR28[4] * Math.Pow(Dp, 2) - N;

            double Disc = Math.Pow(b, 2) - 4 * a * c;
            double x1, x2;
            x1 = (-b + Math.Pow(Disc, 0.5)) / (2 * a);
            x2 = (-b - Math.Pow(Disc, 0.5)) / (2 * a);
            if (x1 > -10 & x1 < 170)
            {
                if (x2 > -10 & x2 < 170)
                {
                    if (x2 > x1) D = x2;
                    else D = x1;
                }
                else D = x1;
            }
            else D = x2;
            return D;
        }

        public double ogr_TG6Qpln(double Qt, double N)
        {
            double D = (N - DSTR27[0] - DSTR27[2] * Qt) / (DSTR27[1] + DSTR27[3] * Qt);
            return D;
        }
        public double ogr_TG6Qplb(double Qt, double N)
        {
            double D = 0;
            double a, b, c;

            a = DSTR26[4] + DSTR26[6] * Qt + DSTR26[8] * Math.Pow(Qt, 2);
            b = DSTR26[1] + DSTR26[3] * Qt + DSTR26[7] * Math.Pow(Qt, 2);
            c = DSTR26[0] + DSTR26[2] * Qt + DSTR26[5] * Math.Pow(Qt, 2) - N;

            double Disc = Math.Pow(b, 2) - 4 * a * c;
            double x1, x2;
            x1 = (-b + Math.Pow(Disc, 0.5)) / (2 * a);
            x2 = (-b - Math.Pow(Disc, 0.5)) / (2 * a);
            if (x1 > x2) D = x1;
            else D = x2;
            return D;
        }
        public double ogr_TG6Qpr(double Dp, double Qt, double N)
        {
            double D = 0;
            double a, b, c;

            a = DSTR28[4] + DSTR28[6] * Qt + DSTR28[8] * Math.Pow(Qt, 2);
            b = DSTR28[1] + DSTR28[3] * Qt + DSTR28[7] * Math.Pow(Qt, 2);
            c = DSTR28[0] + DSTR28[2] * Qt + DSTR28[5] * Math.Pow(Qt, 2) - N;

            double Disc = Math.Pow(b, 2) - 4 * a * c;
            double x1, x2;
            x1 = (-b + Math.Pow(Disc, 0.5)) / (2 * a);
            x2 = (-b - Math.Pow(Disc, 0.5)) / (2 * a);
            if (x1 < x2 & Dp < x1) D = 1001;
            if (x1 > x2 & Dp > x1) D = x1;
            if (x2 < x1 & Dp < x2) D = 1001;
            if (x2 > x1 & Dp > x2) D = x2;
            return D;
        }

        public bool ogr_TG6Qprb(double D, double Qt, double N)
        {
            double a, b, c;

            a = DSTR28[4] + DSTR28[6] * Qt + DSTR28[8] * Math.Pow(Qt, 2);
            b = DSTR28[1] + DSTR28[3] * Qt + DSTR28[7] * Math.Pow(Qt, 2);
            c = DSTR28[0] + DSTR28[2] * Qt + DSTR28[5] * Math.Pow(Qt, 2) - N;

            double Disc = Math.Pow(b, 2) - 4 * a * c;
            double x1, x2;
            x1 = (-b + Math.Pow(Disc, 0.5)) / (2 * a);
            x2 = (-b - Math.Pow(Disc, 0.5)) / (2 * a);
            if ((x1 < x2 & D > x1 - 7 & D < x2 + 7) || (x2 < x1 & D > x2 + 7 & D < x1 - 7)) return false;
            else return true;
        }

        public double D6(double D13, double Qt, double N)
        {
            double D = 0;
            double a, b, c;

            a = DSTR51[4] + DSTR51[11] * Math.Pow(Qt, 2) + DSTR51[13] * Math.Pow(N, 2) + DSTR51[14] * Math.Pow(Qt, 2) * Math.Pow(N, 2) + DSTR51[15] * Qt + DSTR51[19] * N + DSTR51[21] * Qt * N + DSTR51[22] * Math.Pow(Qt, 2) * N + DSTR51[26] * Qt * Math.Pow(N, 2);
            b = DSTR51[1] + DSTR51[7] * Qt + DSTR51[9] * N + DSTR51[10] * Qt * N + DSTR51[16] * Math.Pow(Qt, 2) + DSTR51[20] * Math.Pow(N, 2) + DSTR51[23] * Math.Pow(Qt, 2) * N + DSTR51[24] * Math.Pow(Qt, 2) * Math.Pow(N, 2) + DSTR51[25] * Qt * Math.Pow(N, 2);
            c = DSTR51[0] + DSTR51[2] * Qt + DSTR51[3] * N + DSTR51[5] * Math.Pow(Qt, 2) + DSTR51[6] * Math.Pow(N, 2) + DSTR51[8] * Qt * N + DSTR51[12] * Math.Pow(Qt, 2) * Math.Pow(N, 2) + DSTR51[17] * Math.Pow(Qt, 2) * N + DSTR51[18] * Qt * Math.Pow(N, 2) - D13;

            double Disc = Math.Pow(b, 2) - 4 * a * c;
            double x1, x2;
            x1 = (-b + Math.Pow(Disc, 0.5)) / (2 * a);
            x2 = (-b - Math.Pow(Disc, 0.5)) / (2 * a);
            if (Math.Abs(x1 - D13) < 20) D = x1;
            else D = x2;
            return D;
        }

        public double Nsnto(double N1, double N3, double N4, double N6)
        {
            double sn, sn1, sn2, sn3, sn4, sn5, sn6;
            sn1 = 0; sn2 = 0; sn3 = 0; sn4 = 0; sn5 = 0; sn6 = 0;
            sn1 = -0.0000428233 * Math.Pow(N1, 2) + 0.0052171908 * N1 + 0.0366057976;
            sn2 = 0.0411;
            sn3 = -0.0000311462 * Math.Pow(N3, 2) + 0.0044165282 * N3 + 0.0404727907;
            sn4 = -0.0000134196 * Math.Pow(N4, 2) + 0.0030831795 * N4 + 0.0823025908;
            sn5 = 0.03475;
            sn6 = -0.0000211825 * Math.Pow(N6, 2) + 0.0048747558 * N6 + 0.0797149679;
            sn = 0.49745 + sn1 + sn2 + sn3 + sn4 + sn5 + sn6; // прочие+Т1+Т2+Т3+Т4+Т5+Т6
            return sn;
        } //Расход эл.энергии прочие т/о
        public double Nsnpn()
        {
            double Gekm = (Q2(Dp2) + Q3(P3, Qt3, N3) + Q4(P4, Qt4, N4) + Q5(Dp5) + Q6(Dp6, Qt6, N6)) / 0.595 + D140_13 + D140_30 + D140_45;
            Gekm = Gekm - (Gpvptn(Dptn1) + Gpvptn(Dptn2));
            double npn = (6.5797 * Gekm + 441.4) / 1000;
            return npn;
        }
        public double Gpvptn(double D)
        {
            double G = 0;
            if (D != 0) { G = 14 * D - 830; }
            return G;
        }
        public double DD(double[] M, double[] M1, out double D, out double D1)
        {
            double DD = 0;
            D = 0; D1 = 0;
            for (int i = 0; i < 23; i++)
            {
                D = D + Math.Pow(M[i], 2);
                D1 = D1 + Math.Pow(M1[i], 2);
            }
            D = Math.Pow(D / 23, 0.5);
            D1 = Math.Pow(D1 / 23, 0.5);
            DD = D - D1;
            return DD;
        }

        public int Ogran(double[] Y, int ogrN1, int ogrN3, int ogrN4, int ogrN6, int ogrQt1, int ogrQt3, int ogrQt4, int ogrQt6, out int ogrN31, out int ogrN41, out int ogrN61)
        {
            int ogrN11 = ogrN1; ogrN31 = ogrN3; ogrN41 = ogrN4; ogrN61 = ogrN6;
            if (indtg1 == 1)
            {
                if (ogrQt1 == 1 & ogrN1 > 0) ogrN11 = 6;
                else
                {
                    if (Y[17] == 1)
                    {
                        if (ogrN1 == 1) ogrN11 = 2;
                        if (ogrN1 == 3) ogrN11 = 5;
                    }
                    if (Y[16] == 1 & Y[17] == 2 & ogrN1 > 2) ogrN11 = 6;
                    if (Y[16] == 2 & Y[17] == 2)
                    {
                        if (ogrN1 == 1) ogrN11 = 2;
                        if (ogrN1 > 3) ogrN11 = 6;
                    }
                }
            }
            if (indtg3 == 1)
            {
                if (ogrQt3 == 1 & ogrN3 > 0) ogrN31 = 4;
                else
                {
                    if (Y[19] == 1)
                    {
                        if (ogrN3 > 0) ogrN31 = 3;
                    }
                    if (Y[19] == 2)
                    {
                        if (ogrN3 > 1) ogrN31 = 4;
                    }
                }
            }
            if (indtg4 == 1)
            {
                if (ogrQt4 == 1 & ogrN4 > 0) ogrN41 = 8;
                else
                {
                    if (Y[20] == 1)
                    {
                        if (Y[21] == 1)
                        {
                            if (ogrN4 == 1) ogrN41 = 2;
                            if (ogrN4 == 3) ogrN41 = 6;
                        }
                        if (Y[21] == 2)
                        {
                            if (ogrN4 > 2) ogrN41 = 6;
                            if (ogrN4 > 6) ogrN41 = 8;
                        }
                    }
                    if (Y[20] == 2)
                    {
                        if (Y[21] == 1)
                        {
                            if (ogrN4 == 1) ogrN41 = 2;
                            if (ogrN4 == 3) ogrN41 = 5;
                        }
                        if (Y[21] == 2)
                        {
                            if (ogrN4 == 1) ogrN41 = 2;
                            if (ogrN4 == 4) ogrN41 = 6;
                        }
                        if (ogrN4 > 6) ogrN41 = 8;
                    }
                }
            }
            if (indtg6 == 1)
            {
                if (ogrQt6 == 1 & ogrN6 > 0) ogrN61 = 6;
                else
                {
                    if (Y[23] == 2)
                    {
                        if (ogrN6 == 2 || ogrN6 == 3) ogrN61 = 4;
                        if (ogrN6 == 5) ogrN61 = 6;
                    }
                    if (Y[22] == 1)
                    {
                        if (Y[23] == 1)
                        {
                            if (ogrN6 == 1 || ogrN6 == 2 || ogrN6 == 3) ogrN61 = 4;
                            if (ogrN6 == 6) ogrN61 = 6;
                        }
                    }
                    if (Y[22] == 2)
                    {
                        if (Y[23] == 1)
                        {
                            if (ogrN6 == 1 || ogrN6 == 2) ogrN61 = 3;
                            if (ogrN6 == 5) ogrN61 = 6;
                        }
                    }
                }
            }
            return ogrN11;
        }

        public int Allogr(double Qt1, double Qt3, double Qt4, double Qt6, double Dp6, double N1, double N3, double N4, double N6)
        {
            int n = 0;
            if (Qt1 > 100 || Qt1 < 0 || Qt3 > 92 || Qt3 < 0 || Qt4 < 168 || Qt4 < 0 || Qt6 > 160 || Qt6 < 0 || Dp6 > ogr_TG6bt(Qt6) || Dp6 < 0 || N1 < ogr_TG1l(P1, Qt1) || (Qt1 < ogr_TG1Qt(P1) & N1 > ogr_TG1rb(P1, Qt1)) || (Qt1 > ogr_TG1Qt(P1) & N1 > ogr_TG1rn(P1, Qt1)) || N3 < ogr_TG3l(P3, Qt3) || N3 > ogr_TG3r(P3, Qt3) || N4 < 40 || N4 < ogr_TG4l(P4, Qt4) || N4 > 120 || N4 > ogr_TG4r(P4, Qt4) || (Qt6 < ogr_TG6Dp(Dp6) & N6 < ogr_TG6ln(Dp6, Qt6)) || (Qt6 > ogr_TG6Dp(Dp6) & N6 < ogr_TG6lb(Dp6, Qt6)) || N6 > ogr_TG6r(Dp6, Qt6)) n = 1;
            return n;
        }

        public double Fin(out double S, out double Nsn, out double Kotrt, out double Kotrk, out double nnet, out double ntp, out double q)
        {
            double B;
            S = N1 + NN2(Dp2) + N3 + N4 + NN5(Dp5) + N6;
            dQe = (1 - kp) * (Dp2 * k13 + Dp5 * k30 + Dp6 * k45) + (1 - kt) * (Qt1 + Qt3 + Qt4 + Qt6) * 0.54; //прирост тепла на производство электроэнергии
            double q2 = 0.1536 * Dp2 - 9.4561; double q5 = 0.1077 * Dp5 - 9.7618;
            if (Dp2 == 0) q2 = 0; if (Dp5 == 0) q5 = 0;
            Qe = (D1(P1, Qt1, N1) * 0.56 - Qt1) + q2 + (Q3(P3, Qt3, N3) - Qt3) + (Q4(P4, Qt4, N4) - Qt4) + q5 + (Q6(Dp6, Qt6, N6) - Qt6 - Qp6(Dp6, Qt6, N6)); // затраты тепла на выработку э/э
            q = Qe / S * 1000;
            Qsnto = 0.6 * Qe / 100; // расход тепла на СН ТО
            Kotrt = 1 + dQe / (Qe + Qsnto); // Котрт
            Qsnko = 0.005 * Gek * 0.595; // расход тепла на СН КО                    
            ntp = 100 - (1.5 * Qknom / (Gek * 0.595)); // КПД теплового потока
            Kotrk = 1 + dQe / ((Gek * 0.595 - Qsnko) * (ntp / 100)); // Котрk
            anas = 0.86 * 0.96 * (Nsnsw(Gsetw) / 1000 * Gsetw) / Qs / 100;
            apvk = Qpvk / Qs * 100;
            Ke = (Qe + Qsnto) / (Qe + Qsnto + Qs * (100 - anas - apvk - 0) * 0.01);
            Nko = Nsnpn() + Nsntd(G1p, G2p, G3p, G4p, G5p, G6p, G7p) + Nsnko(Gek) + Nsnpod(Ghob, Gdist, Gpts) + NsntdWB(Qpvk); // Расход э/э на С.Н. котельного отделения
            Ncn = Ncnn();
            Nto = Ncn + Nsnto(N1, N3, N4, N6);
            Nsnee = Ke * Nko + Nto;
            Nsn = (Nto + Nko + Nsnsw(Gsetw));
            nnet = (nk * (100 - 0.5) * (100 - Nsnee / S * 100)) / (100 - Nto / S * 100) / 100;
            B = (Gek * 0.595) / (7 * ntp / 100 * nk / 100) + Qptvm(Qpvk) / (7 * ntp / 100); // Расход топлива
            return B;
        }

        public double Nsnsw(double G)
        {
            double nsw = 0.0006452 * G + 0.3233907;
            return nsw;
        }
        public double Nsntd(double G1, double G2, double G3, double G4, double G5, double G6, double G7)
        {
            double ntd, ntd1, ntd2, ntd3, ntd4, ntd5, ntd6, ntd7;
            ntd1 = 0;
            if (G1 > 0)
            {
                if (G1 <= G1max) ntd1 = 12.2847385406494 - 0.0368681363761425 * G1 + 5.8394820371177E-5 * Math.Pow(G1, 2);
                if (G1 > G1max) ntd1 = 21.0949935913086 - 0.0579417012631893 * G1 + 6.21053040958941E-5 * Math.Pow(G1, 2);
            }
            ntd2 = 0;
            if (G2 > 0)
            {
                if (G2 <= G2max) ntd2 = 10.1671981811523 - 0.0245829522609711 * G2 + 3.57363969669677E-5 * Math.Pow(G2, 2);
                if (G2 > G2max) ntd2 = 20.9693126678467 - 0.0643438026309013 * G2 + 7.23057819413953E-5 * Math.Pow(G2, 2);
            }
            ntd3 = 0;
            if (G3 > 0)
            {
                if (G3 <= G3max) ntd3 = 10.1671981811523 - 0.0245829522609711 * G3 + 3.57363969669677E-5 * Math.Pow(G3, 2);
                if (G3 > G3max) ntd3 = 20.9693126678467 - 0.0643438026309013 * G3 + 7.23057819413953E-5 * Math.Pow(G3, 2);
            }
            ntd4 = 0;
            if (G4 > 0)
            {
                if (G4 <= G4max) ntd4 = 12.2847385406494 - 0.0368681363761425 * G4 + 5.8394820371177E-5 * Math.Pow(G4, 2);
                if (G4 > G4max) ntd4 = 21.0949935913086 - 0.0579417012631893 * G4 + 6.21053040958941E-5 * Math.Pow(G4, 2);
            }
            ntd5 = 0;
            if (G5 > 0)
            {
                if (G5 <= G5max) ntd5 = 7.24247455596924 - 0.00974651612341404 * G5 + 6.62923366689938E-6 * Math.Pow(G5, 2);
                if (G5 > G5max) ntd5 = 3.46760940551758 + 0.0051411404274404 * G5 - 1.53174269712508E-7 * Math.Pow(G5, 2);
            }
            ntd6 = 0;
            if (G6 > 0)
            {
                if (G6 <= G6max) ntd6 = 3.58525681495667 + 0.0115508949384093 * G6 - 1.48867038660683E-5 * Math.Pow(G6, 2);
                if (G6 > G6max) ntd6 = 0.400030344724655 + 0.0199999213218689 * G6;
            }
            ntd7 = 0;
            if (G7 > 0)
            {
                if (G7 <= G7max) ntd7 = 3.58525681495667 + 0.0115508949384093 * G7 - 1.48867038660683E-5 * Math.Pow(G7, 2);
                if (G7 > G7max) ntd7 = 0.400030344724655 + 0.0199999213218689 * G7;
            }
            ntd = (ntd1 * G1 + ntd2 * G2 + ntd3 * G3 + ntd4 * G4 + ntd5 * G5 + ntd6 * G6 + ntd7 * G7) * 0.595 / 1000;
            return ntd;
        }
        public double Nsnko(double G)
        {
            double snk;
            double Q = G * 0.595;
            snk = 7.83572912216186 - 0.0300981141626835 * Q + 5.86559108342044E-5 * Math.Pow(Q, 2) - 5.79804009248619E-8 * Math.Pow(Q, 3) + 2.80739320679402E-11 * Math.Pow(Q, 4) - 5.28290637554244E-15 * Math.Pow(Q, 5);
            return snk;
        } // Расход эл.энергии прочие к/о
        public double Nsnpod(double Ghob, double Gdist, double Gpts)
        {
            double nhob = 33.8383102416992 - 0.268439769744873 * Ghob + 0.00102928199339658 * Math.Pow(Ghob, 2) - 2.11662700166926E-6 * Math.Pow(Ghob, 3) + 2.40049424782285E-9 * Math.Pow(Ghob, 4) - 1.41589246837465E-12 * Math.Pow(Ghob, 5) + 3.39281773410009E-16 * Math.Pow(Ghob, 6);
            double ndist = 0;
            if (Gdist < 150) ndist = 12.6100244522095 - 0.128100603818893 * Gdist + 0.000418003095546737 * Math.Pow(Gdist, 2);
            if (Gdist >= 150) ndist = 6.00028562545776 - 0.0202885624021292 * Gdist + 2.48571068368619E-5 * Math.Pow(Gdist, 2);
            double npts = 40.664680480957 - 0.428236991167068 * Gpts + 0.00161372788716108 * Math.Pow(Gpts, 2) - 4.02618951511613E-7 * Math.Pow(Gpts, 3) - 1.11690772186535E-8 * Math.Pow(Gpts, 4) + 2.46290592087473E-11 * Math.Pow(Gpts, 5) - 1.59311040921761E-14 * Math.Pow(Gpts, 6);
            double npod = (nhob * Ghob + ndist * Gdist + npts * Gpts) / 1000;
            return npod;
        }
        public double Ncnn()
        {
            double Ncn, Ncn1, Ncn3, Ncn4, Ncn6;
            Ncn1 = 0; Ncn3 = 0; Ncn4 = 0; Ncn6 = 0;
            if (indtg1 == 1)
            {
                if (indcn1 == 1) Ncn1 = 310;
                if (indcn1 == 2) Ncn1 = 620;
            }
            if (indtg3 == 1)
            {
                if (indcn3 == 1) Ncn3 = 310;
                if (indcn3 == 2) Ncn3 = 620;
            }
            if (indtg4 == 1)
            {
                if (indcn4 == 1) Ncn4 = 620;
                if (indcn4 == 2) Ncn4 = 1240;
            }
            if (indtg6 == 1)
            {
                Ncn6 = 1100;
            }
            Ncn = (Ncn1 + Ncn3 + Ncn4 + Ncn6) / 1000;
            return Ncn;
        }
        public double nptvm(double Q)
        {
            double n = 0;
            double[] Xp = { 1, Q };
            for (int i = 0; i < 2; i++)
            {
                n = n + Xp[i] * DSTR47[i];
            }
            return n;
        }

        public double nek(double G, int ind, out double G1, out double G2, out double G3, out double G4, out double G5, out double G6, out double G7)
        {
            double n; double z = 0;            
            double[] K1 = { 0, 0 }; double[] K2 = { 0, 0 }; double[] K3 = { 0, 0 }; double[] K4 = { 0, 0 }; double[] K5 = { 0, 0 }; double[] K6 = { 0, 0 }; double[] K7 = { 0, 0 };
            G1 = 0; G2 = 0; G3 = 0; G4 = 0; G5 = 0; G6 = 0; G7 = 0;
            double G11, G21, G31, G41, G51, G61, G71;
            double Gmin = 0; double Gmax = 0;

            if (G < 200) { n = 0; return n; }

            if (indk1 == 1) { z++; Gmin = Gmin + G1min; Gmax = Gmax + G1max; } if (indk2 == 1) { z++; Gmin = Gmin + G2min; Gmax = Gmax + G2max; } if (indk3 == 1) { z++; Gmin = Gmin + G3min; Gmax = Gmax + G3max; } if (indk4 == 1) { z++; Gmin = Gmin + G4min; Gmax = Gmax + G4max; }
            if (indk5 == 1) { z++; Gmin = Gmin + G5min; Gmax = Gmax + G5max; } if (indk6 == 1) { z++; Gmin = Gmin + G6min; Gmax = Gmax + G6max; } if (indk7 == 1) { z++; Gmin = Gmin + G7min; Gmax = Gmax + G7max; }

            if ((G < Gmin || G > Gmax) & combobox20.Text == "запрещено" & combobox10.Text == "неактивен") { return 0; }

            while ((G > Gmax) & (combobox20.Text == "разрешено" || combobox10.Text == "активен"))
            {
                if (indk3 == 0 & combobox13.Text != "ремонт" & G > Gmax) { indk3 = 1; Gmax = Gmax + G3max; Gmin = Gmin + G3min; z++; }
                if (indk2 == 0 & combobox14.Text != "ремонт" & G > Gmax) { indk2 = 1; Gmax = Gmax + G2max; Gmin = Gmin + G2min; z++; }
                if (indk6 == 0 & combobox15.Text != "ремонт" & G > Gmax) { indk6 = 1; Gmax = Gmax + G6max; Gmin = Gmin + G6min; z++; }
                if (indk5 == 0 & combobox16.Text != "ремонт" & G > Gmax) { indk5 = 1; Gmax = Gmax + G5max; Gmin = Gmin + G5min; z++; }
                if (indk7 == 0 & combobox17.Text != "ремонт" & G > Gmax) { indk7 = 1; Gmax = Gmax + G7max; Gmin = Gmin + G7min; z++; }
                if (indk4 == 0 & combobox18.Text != "ремонт" & G > Gmax) { indk4 = 1; Gmax = Gmax + G4max; Gmin = Gmin + G4min; z++; }
                if (indk1 == 0 & combobox19.Text != "ремонт" & G > Gmax) { indk1 = 1; Gmax = Gmax + G6max; Gmin = Gmin + G1min; z++; }
            }

            while (G < Gmin & (combobox20.Text == "разрешено" || combobox10.Text == "активен"))
            {
                if (indk1 == 1) { indk1 = 0; Gmin = Gmin - G1min; z--; continue; }
                if (indk4 == 1) { indk4 = 0; Gmin = Gmin - G4min; z--; continue; }
                if (indk5 == 1) { indk5 = 0; Gmin = Gmin - G5min; z--; continue; }
                if ((indk7 == 1 & indtg6 == 1 & indk6 == 1) || (indk7 == 1 & indtg6 == 0)) { indk7 = 0; Gmin = Gmin - G7min; z--; continue; }
                if ((indk6 == 1 & indtg6 == 1 & indk7 == 1) || (indk6 == 1 & indtg6 == 0)) { indk6 = 0; Gmin = Gmin - G6min; z--; continue; }
                if (indk2 == 1) { indk2 = 0; Gmin = Gmin - G2min; z--; continue; }
                if (indk3 == 1) { indk3 = 0; Gmin = Gmin - G3min; z--; continue; }
            }

            if (indk1 == 1) K1[1] = DSTR37[1]; if (indk2 == 1) K2[1] = DSTR38[1]; if (indk3 == 1) K3[1] = DSTR38[1];
            if (indk4 == 1) K4[1] = DSTR37[1]; if (indk5 == 1) K5[1] = DSTR39[1]; if (indk6 == 1) K6[1] = DSTR40[1]; if (indk7 == 1) K7[1] = DSTR40[1];

            double z1 = z; double G0 = G;
            do
            {
                if (indk1 == 1)
                {
                    G1 = G1 + G0 / z1;
                    if (G1 < G1min & G1 != 0) { G1 = G1min; } if (G1 > G1max & G1 != 0) { G1 = G1max; }
                }
                if (indk2 == 1)
                {
                    G2 = G2 + G0 / z1;
                    if (G2 < G2min & G2 != 0) { G2 = G2min; } if (G2 > G2max & G2 != 0) { G2 = G2max; }
                }
                if (indk3 == 1)
                {
                    G3 = G3 + G0 / z1;
                    if (G3 < G3min & G3 != 0) { G3 = G3min; } if (G3 > G3max & G3 != 0) { G3 = G3max; }
                }
                if (indk4 == 1)
                {
                    G4 = G4 + G0 / z1;
                    if (G4 < G4min & G4 != 0) { G4 = G4min; } if (G4 > G4max & G4 != 0) { G4 = G4max; }
                }
                if (indk5 == 1)
                {
                    G5 = G5 + G0 / z1;
                    if (G5 < G5min & G5 != 0) { G5 = G5min; } if (G5 > G5max & G5 != 0) { G5 = G5max; }
                }
                if (indk6 == 1)
                {
                    G6 = G6 + G0 / z1;
                    if (G6 < G6min & G6 != 0) { G6 = G6min; } if (G6 > G6max & G6 != 0) { G6 = G6max; }
                }
                if (indk7 == 1)
                {
                    G7 = G7 + G0 / z1;
                    if (G7 < G7min & G7 != 0) { G7 = G7min;  } if (G7 > G7max & G7 != 0) { G7 = G7max; }
                }
                G0 = G - G1 - G2 - G3 - G4 - G5 - G6 - G7;
            } while (G0 > 0.1);
            int k1, k2, k3, k4, k5, k6, k7; // индикатор достижения расхода котлом максимального или минимального значения
            k1 = 0; k2 = 0; k3 = 0; k4 = 0; k5 = 0; k6 = 0; k7 = 0;
            do
            {
                G11 = G1; G21 = G2; G31 = G3; G41 = G4; G51 = G5; G61 = G6; G71 = G7;
                if ((G1 == G1min || G1 == G1max) & k1 == 0) { K1[1] = 0; z--; k1 = 1; } if ((G2 == G2min || G2 == G2max) & k2 == 0) { K2[1] = 0; z--; k2 = 1; }
                if ((G3 == G3min || G3 == G3max) & k3 == 0) { K3[1] = 0; z--; k3 = 1; } if ((G4 == G4min || G4 == G4max) & k4 == 0) { K4[1] = 0; z--; k4 = 1; }
                if ((G5 == G5min || G5 == G5max) & k5 == 0) { K5[1] = 0; z--; k5 = 1; } if ((G6 == G6min || G6 == G6max) & k6 == 0) { K6[1] = 0; z--; k6 = 1; }
                if ((G7 == G7min || G7 == G7max) & k7 == 0) { K7[1] = 0; z--; k7 = 1; }
                double df = (K1[1] + K2[1] + K3[1] + K4[1] + K5[1] + K6[1] + K7[1]) / z;
                if (indk1 == 1 & G1 != G1min & G1 != G1max) G1 = G1 - 100 * (K1[1] - df);
                if (indk2 == 1 & G2 != G2min & G2 != G2max) G2 = G2 - 100 * (K2[1] - df);
                if (indk3 == 1 & G3 != G3min & G3 != G3max) G3 = G3 - 100 * (K3[1] - df);
                if (indk4 == 1 & G4 != G4min & G4 != G4max) G4 = G4 - 100 * (K4[1] - df);
                if (indk5 == 1 & G5 != G5min & G5 != G5max) G5 = G5 - 100 * (K5[1] - df);
                if (indk6 == 1 & G6 != G6min & G6 != G6max) G6 = G6 - 100 * (K6[1] - df);
                if (indk7 == 1 & G7 != G7min & G7 != G7max) G7 = G7 - 100 * (K7[1] - df);
                if (indk1 == 1) { if (G1 < G1min) G1 = G1min; if (G1 > G1max) G1 = G1max; }
                if (indk2 == 1) { if (G2 < G2min) G2 = G2min; if (G2 > G2max) G2 = G2max; }
                if (indk3 == 1) { if (G3 < G3min) G3 = G3min; if (G3 > G3max) G3 = G3max; }
                if (indk4 == 1) { if (G4 < G4min) G4 = G4min; if (G4 > G4max) G4 = G4max; }
                if (indk5 == 1) { if (G5 < G5min) G5 = G5min; if (G5 > G5max) G5 = G5max; }
                if (indk6 == 1) { if (G6 < G6min) G6 = G6min; if (G6 > G6max) G6 = G6max; }
                if (indk7 == 1) { if (G7 < G7min) G7 = G7min; if (G7 > G7max) G7 = G7max; }

                double dfg = G1 + G2 + G3 + G4 + G5 + G6 + G7;

            } while ((G11 != G1 & (G1 != G1min & G1 != G1max)) || (G21 != G2 & (G2 != G2min & G2 != G2max)) || (G31 != G3 & (G3 != G3min & G3 != G3max)) || (G41 != G4 & (G4 != G4min & G4 != G4max)) || (G51 != G5 & (G5 != G5min & G5 != G5max)) || (G61 != G6 & (G6 != G6min & G6 != G6max)) || (G71 != G7 & (G7 != G7min & G7 != G7max)));

            double[] X1 = { 1, G1 }; double[] X2 = { 1, G2 }; double[] X3 = { 1, G3 }; double[] X4 = { 1, G4 };
            double[] X5 = { 1, G5 }; double[] X6 = { 1, G6 }; double[] X7 = { 1, G7 };
            if (indk1 != 1) { X1[0] = 0; X1[1] = 0; } if (indk2 != 1) { X2[0] = 0; X2[1] = 0; } if (indk3 != 1) { X3[0] = 0; X3[1] = 0; }
            if (indk4 != 1) { X4[0] = 0; X4[1] = 0; } if (indk5 != 1) { X5[0] = 0; X5[1] = 0; } if (indk6 != 1) { X6[0] = 0; X6[1] = 0; }
            if (indk7 != 1) { X7[0] = 0; X7[1] = 0; }

            n = 0;
            for (int i = 0; i < 2; i++)
            {
                n = n + X1[i] * DSTR37[i] + X2[i] * DSTR38[i] + X3[i] * DSTR38[i] + X4[i] * DSTR37[i] + X5[i] * DSTR39[i] + X6[i] * DSTR40[i] + X7[i] * DSTR40[i];
            }
            n = G * 0.595 / (7 * n) * 100;

            Qknom = 0;
            if (indk1 == 1) Qknom = Qknom + 420 * 0.595; if (indk2 == 1) Qknom = Qknom + 420 * 0.595; if (indk3 == 1) Qknom = Qknom + 420 * 0.595; if (indk4 == 1) Qknom = Qknom + 420 * 0.595;
            if (indk5 == 1) Qknom = Qknom + 500 * 0.595; if (indk6 == 1) Qknom = Qknom + 400 * 0.595; if (indk7 == 1) Qknom = Qknom + 400 * 0.595; // номинальная тепловая нагрузка

            if (ind == 0)
            {
                if (combobox10.Text == "активен") { indk1 = 0; indk2 = 0; indk3 = 0; indk4 = 0; indk5 = 0; indk6 = 0; indk7 = 0; }
                if (combobox20.Text == "разрешено")
                {
                    if (combobox13.Text != "работа") indk1 = 0; if (combobox14.Text != "работа") indk2 = 0; if (combobox15.Text != "работа") indk3 = 0; if (combobox16.Text != "работа") indk4 = 0;
                    if (combobox17.Text != "работа") indk5 = 0; if (combobox18.Text != "работа") indk6 = 0; if (combobox19.Text != "работа") indk7 = 0;
                }
            }
            return n;
        }

        public double[,] ishd(int ind1, int ind2)
        {
            string str; double str1;
            double[,] X = new double[32, 32];
            excelapp = new Excel.Application();
            excelapp.Visible = true;
            excelappworkbooks = excelapp.Workbooks;
            //Открываем книгу и получаем на нее ссылку           
            excelappworkbook = excelapp.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + @"\Макет.xlsx",
             Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
             Type.Missing, Type.Missing, Type.Missing);
            excelsheets = excelappworkbook.Worksheets;
            //Получаем ссылку на лист 1
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);

            for (int m = 0; m < 31; m++)
            {
                for (int n = ind1; n < ind2 + 1; n++)
                {
                    excelcells = (Excel.Range)excelworksheet.Cells[m + 3, n + 2];
                    str = Convert.ToString(excelcells.Value2);
                    double.TryParse(str, out str1);
                    X[m, n] = str1;
                }
            }
            excelappworkbook.Close(true, Type.Missing, Type.Missing);
            excelapp.Quit();
            GC.Collect(); // убрать за собой
            return X;
        }

        public void input(double[,] X, int t, int tex)
        {
            double[] Xmest = new double[33];
            Xmest[0] = X[t, 0]; Xmest[1] = X[t, 1]; Xmest[2] = X[t, 2]; Xmest[3] = X[t, 3]; Xmest[4] = X[t, 4]; Xmest[5] = X[t, 5]; Xmest[6] = X[t, 6]; Xmest[7] = X[t, 7]; Xmest[8] = NN5(X[t, 8]); Xmest[9] = X[t, 8]; Xmest[10] = X[t, 11]; Xmest[11] = X[t, 10];
            Xmest[12] = X[t, 9]; Xmest[13] = X[t, 14]; Xmest[14] = X[t, 15]; Xmest[15] = X[t, 16]; Xmest[16] = X[t, 17]; Xmest[17] = X[t, 18]; Xmest[18] = X[t, 19]; Xmest[19] = X[t, 20]; Xmest[20] = X[t, 32]; Xmest[21] = X[t, 30]; Xmest[22] = X[t, 31]; Xmest[23] = X[t, 73];
            Xmest[24] = X[t, 35]; Xmest[25] = X[t, 36]; Xmest[26] = X[t, 42];
            if (iii11 != "") Xmest[27] = Convert.ToDouble(iii11);
            if (iii21 != "") Xmest[28] = Convert.ToDouble(iii21);
            if (iii31 != "") Xmest[29] = Convert.ToDouble(iii31);
            if (iii41 != "") Xmest[30] = Convert.ToDouble(iii41);
            if (iii51 != "") Xmest[31] = Convert.ToDouble(iii51);
            if (iii61 != "") Xmest[32] = Convert.ToDouble(iii61);
            //Получаем ссылку на лист 1
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
            for (int n = 0; n < 27; n++)
            {
                excelcells = (Excel.Range)excelworksheet.Cells[n + 3, tex + 2];                
                excelcells.Value = Xmest[n];                
            }
            excelcells = (Excel.Range)excelworksheet.Cells[28 + 3, tex + 2];
            excelcells.Value = iii11 + iii21 + iii31 + iii41 + iii51 + iii61;
        }

        public void output(double[,] X, int t)
        {
            label96.Content = Math.Round(X[t, 9], 0).ToString(); label87.Content = Math.Round(X[t, 10], 0).ToString(); label95.Content = Math.Round(X[t, 11], 0).ToString(); label94.Content = Math.Round(NN5(X[t, 8]), 0).ToString();
            label86.Content = Math.Round(X[t, 7], 0).ToString(); label93.Content = Math.Round(X[t, 6], 0).ToString(); label85.Content = Math.Round(X[t, 5], 0).ToString(); label92.Content = Math.Round(X[t, 4], 0).ToString();
            label91.Content = Math.Round(X[t, 2], 0).ToString(); label99.Content = Math.Round(X[t, 12], 0).ToString(); label89.Content = Math.Round(X[t, 13], 0).ToString(); label27.Content = Math.Round(X[t, 14], 0).ToString();
            label31.Content = Math.Round(X[t, 15], 0).ToString(); label56.Content = Math.Round(X[t, 16], 0).ToString(); label59.Content = Math.Round(X[t, 17], 0).ToString(); label60.Content = Math.Round(X[t, 18], 0).ToString();
            label63.Content = Math.Round(X[t, 19], 0).ToString(); label64.Content = Math.Round(X[t, 20], 0).ToString(); label79.Content = Math.Round(X[t, 21], 0).ToString(); label80.Content = Math.Round(X[t, 22], 0).ToString(); label81.Content = Math.Round(X[t, 23], 0).ToString(); label82.Content = Math.Round(X[t, 24], 0).ToString();
            label83.Content = Math.Round(X[t, 25], 0).ToString(); label84.Content = Math.Round(X[t, 3], 0).ToString(); label88.Content = Math.Round(X[t, 26], 0).ToString(); label97.Content = Math.Round(X[t, 27], 0).ToString();
            label98.Content = Math.Round(X[t, 28], 0).ToString(); label100.Content = Math.Round(X[t, 29], 0).ToString(); label119.Content = Math.Round(X[t, 30], 1).ToString(); label120.Content = Math.Round(X[t, 31], 1).ToString(); label121.Content = Math.Round(X[t, 32], 1).ToString();
            label122.Content = Math.Round(X[t, 33], 1).ToString(); label123.Content = Math.Round(X[t, 34], 2).ToString(); label124.Content = Math.Round(X[t, 35], 1).ToString(); label125.Content = Math.Round(X[t, 36], 1).ToString(); label126.Content = Math.Round(X[t, 37], 0).ToString(); label129.Content = Math.Round(X[t, 73], 3).ToString(); label132.Content = Math.Round(X[t, 39], 0).ToString();
            label118.Content = Math.Round(X[t, 40], 1).ToString(); label106.Content = Math.Round(X[t, 8], 0).ToString(); label117.Content = Math.Round(X[t, 41], 0).ToString();
            textBox6.Text = Math.Round(X[t, 42], 0).ToString(); label101.Content = Math.Round(X[t, 26], 1).ToString(); label102.Content = Math.Round(X[t, 44], 1).ToString(); label103.Content = Math.Round(X[t, 46], 1).ToString();
            label135.Content = Math.Round(X[t, 1], 0).ToString(); label136.Content = Math.Round(X[t, 45], 0).ToString(); label137.Content = Math.Round(X[t, 0], 0).ToString();
            label138.Content = Math.Round(X[t, 47], 1).ToString(); label139.Content = Math.Round(X[t, 48], 1).ToString(); label140.Content = Math.Round(X[t, 49], 0).ToString();
            if (X3[t, 63] == 1) radioButton11.IsChecked = true; if (X3[t, 63] == 2) radioButton14.IsChecked = true; if (X3[t, 63] == 0) { radioButton11.IsChecked = false; radioButton14.IsChecked = false; }
            if (X3[t, 64] == 1) radioButton4.IsChecked = true; if (X3[t, 64] == 2) radioButton3.IsChecked = true; if (X3[t, 64] == 0) { radioButton4.IsChecked = false; radioButton3.IsChecked = false; }
            if (X3[t, 65] == 1) radioButton5.IsChecked = true; if (X3[t, 65] == 2) radioButton6.IsChecked = true; if (X3[t, 65] == 0) { radioButton5.IsChecked = false; radioButton6.IsChecked = false; }
            if (X3[t, 66] == 1) radioButton1.IsChecked = true; if (X3[t, 66] == 2) radioButton2.IsChecked = true; if (X3[t, 66] == 0) { radioButton1.IsChecked = false; radioButton2.IsChecked = false; }
            if (X3[t, 67] == 1) radioButton8.IsChecked = true; if (X3[t, 67] == 2) radioButton10.IsChecked = true; if (X3[t, 67] == 0) { radioButton8.IsChecked = false; radioButton10.IsChecked = false; }
            if (X3[t, 68] == 1) radioButton13.IsChecked = true; if (X3[t, 68] == 2) radioButton12.IsChecked = true; if (X3[t, 68] == 0) { radioButton13.IsChecked = false; radioButton12.IsChecked = false; }
            textBox13.Text = Math.Round(X3[t, 69]).ToString(); textBox22.Text = Math.Round(X3[t, 70]).ToString(); textBox19.Text = Math.Round(X3[t, 71]).ToString(); textBox20.Text = Math.Round(X3[t, 72]).ToString();
            textBox53.Text = Math.Round(X[t, 74], 2).ToString(); textBox54.Text = Math.Round(X[t, 75], 2).ToString(); textBox55.Text = Math.Round(X[t, 76], 2).ToString(); textBox56.Text = Math.Round(X[t, 77], 2).ToString();
            label1.Content = Math.Round(Qpodp, 1).ToString(); label2.Content = Math.Round(X[t, 14] + X[t, 15] + X[t, 16] + X[t, 17] + X[t, 18] + X[t, 19] + X[t, 20], 0).ToString();
        }

        public double tcwout(double D1, double D3, double D4, double D6, double G1, double G3, double G4, double G6, double t)
        {
            double tt = 0;
            double e = -0.5913;
            double f = 598.49;
            double A, B, C, D, E, F1, F3, F4, F6;
            A = DSTR41[0] + DSTR41[1] * t;
            B = DSTR41[2] + DSTR41[3] * t;
            C = 0; F1 = 0; F3 = 0; F4 = 0; F6 = 0; E = 0;
            if (G1 != 0) C = D1 / (1 - e * D1 / G1);
            if (G3 != 0) C = C + D3 / (1 - e * D3 / G3);
            if (G4 != 0) C = C + D4 / (1 - e * D4 / G4);
            if (G6 != 0) C = C + D6 / (1 - e * D6 / G6);
            D = G1 + G3 + G4 + G6;
            if (G1 != 0) E = D1 * dtk1 / (1 - e * D1 / G1);
            if (G3 != 0) E = E + D3 * dtk3 / (1 - e * D3 / G3);
            if (G4 != 0) E = E + D4 * dtk4 / (1 - e * D4 / G4);
            if (G6 != 0) E = E + D6 * dtk6 / (1 - e * D6 / G6);
            if (indtg1 == 1) F1 = D1 * G1 / (G1 - e * D1);
            if (indtg3 == 1) F3 = D3 * G3 / (G3 - e * D3);
            if (indtg4 == 1) F4 = D4 * G4 / (G4 - e * D4);
            if (indtg6 == 1) F6 = D6 * G6 / (G6 - e * D6);
            if (indtg1 != 0 || indtg3 != 0 || indtg4 != 0 || indtg6 != 0) tt = (((A * D + B * e * E + B * f * C) / (D - B * e * C)) * (D + e * (F1 + F3 + F4 + F6)) + ((e * dtk1 + f) * F1 + (e * dtk3 + f) * F3 + (e * dtk4 + f) * F4 + (e * dtk6 + f) * F6)) / D;
            else tt = t;
            if (tt < 10) tt = 10;
            return tt;
        }

        private void combobox21_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (combobox21.SelectedItem == null) return;
            string cmb = combobox21.SelectedItem.ToString();
            cmb = cmb.Substring(0, 3);
            int cmb1;
            int.TryParse(cmb, out cmb1);
            cmb1--;
            if (X3[cmb1, 50] == 0 & combobox1.Text != "ремонт") combobox1.Text = "резерв";
            if (X3[cmb1, 50] == 1 & combobox1.Text != "ремонт") combobox1.Text = "работа";
            if (X3[cmb1, 51] == 0 & combobox2.Text != "ремонт") combobox2.Text = "резерв";
            if (X3[cmb1, 51] == 1 & combobox2.Text != "ремонт") combobox2.Text = "работа";
            if (X3[cmb1, 52] == 0 & combobox3.Text != "ремонт") combobox3.Text = "резерв";
            if (X3[cmb1, 52] == 1 & combobox3.Text != "ремонт") combobox3.Text = "работа";
            if (X3[cmb1, 53] == 0 & combobox4.Text != "ремонт") combobox4.Text = "резерв";
            if (X3[cmb1, 53] == 1 & combobox4.Text != "ремонт") combobox4.Text = "работа";
            if (X3[cmb1, 54] == 0 & combobox5.Text != "ремонт") combobox5.Text = "резерв";
            if (X3[cmb1, 54] == 1 & combobox5.Text != "ремонт") combobox5.Text = "работа";
            if (X3[cmb1, 55] == 0 & combobox6.Text != "ремонт") combobox6.Text = "резерв";
            if (X3[cmb1, 55] == 1 & combobox6.Text != "ремонт") combobox6.Text = "работа";
            if (X3[cmb1, 56] == 0 & combobox13.Text != "ремонт") combobox13.Text = "резерв";
            if (X3[cmb1, 56] == 1 & combobox13.Text != "ремонт") combobox13.Text = "работа";
            if (X3[cmb1, 57] == 0 & combobox14.Text != "ремонт") combobox14.Text = "резерв";
            if (X3[cmb1, 57] == 1 & combobox14.Text != "ремонт") combobox14.Text = "работа";
            if (X3[cmb1, 58] == 0 & combobox15.Text != "ремонт") combobox15.Text = "резерв";
            if (X3[cmb1, 58] == 1 & combobox15.Text != "ремонт") combobox15.Text = "работа";
            if (X3[cmb1, 59] == 0 & combobox16.Text != "ремонт") combobox16.Text = "резерв";
            if (X3[cmb1, 59] == 1 & combobox16.Text != "ремонт") combobox16.Text = "работа";
            if (X3[cmb1, 60] == 0 & combobox17.Text != "ремонт") combobox17.Text = "резерв";
            if (X3[cmb1, 60] == 1 & combobox17.Text != "ремонт") combobox17.Text = "работа";
            if (X3[cmb1, 61] == 0 & combobox18.Text != "ремонт") combobox18.Text = "резерв";
            if (X3[cmb1, 61] == 1 & combobox18.Text != "ремонт") combobox18.Text = "работа";
            if (X3[cmb1, 62] == 0 & combobox19.Text != "ремонт") combobox19.Text = "резерв";
            if (X3[cmb1, 62] == 1 & combobox19.Text != "ремонт") combobox19.Text = "работа";
            if (X3[cmb1, 47] == 0)
            {
                if (combobox7.Text != "ремонт") combobox7.Text = "резерв";
            }
            else combobox7.Text = "работа";
            if (X3[cmb1, 48] == 0)
            {
                if (combobox8.Text != "ремонт") combobox8.Text = "резерв";
            }
            else combobox8.Text = "работа";
            output(X3, cmb1);
        }

        private void WindowClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            excelapp.Quit();
        }

        public void Stream()
        {
            StreamReader streamReader;
            STR1 = new string[8]; streamReader = new StreamReader("ТГ-1.txt"); STR1 = File.ReadAllLines("ТГ-1.txt");
            DSTR1 = new double[8]; DSTR1 = STR1.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR2 = new string[27]; streamReader = new StreamReader("ТГ-3(1).txt"); STR2 = File.ReadAllLines("ТГ-3(1).txt");
            DSTR2 = new double[27]; DSTR2 = STR2.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR3 = new string[27]; streamReader = new StreamReader("ТГ-3(2).txt"); STR3 = File.ReadAllLines("ТГ-3(2).txt");
            DSTR3 = new double[27]; DSTR3 = STR3.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR4 = new string[27]; streamReader = new StreamReader("ТГ-4(1).txt"); STR4 = File.ReadAllLines("ТГ-4(1).txt");
            DSTR4 = new double[27]; DSTR4 = STR4.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR5 = new string[27]; streamReader = new StreamReader("ТГ-4(2).txt"); STR5 = File.ReadAllLines("ТГ-4(2).txt");
            DSTR5 = new double[27]; DSTR5 = STR5.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR6 = new string[27]; streamReader = new StreamReader("ТГ-6(1).txt"); STR6 = File.ReadAllLines("ТГ-6(1).txt");
            DSTR6 = new double[27]; DSTR6 = STR6.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR7 = new string[27]; streamReader = new StreamReader("ТГ-6(2).txt"); STR7 = File.ReadAllLines("ТГ-6(2).txt");
            DSTR7 = new double[27]; DSTR7 = STR7.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR8 = new string[8]; streamReader = new StreamReader("расход в конденсатор ТГ-1.txt"); STR8 = File.ReadAllLines("расход в конденсатор ТГ-1.txt");
            DSTR8 = new double[8]; DSTR8 = STR8.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR9 = new string[8]; streamReader = new StreamReader("расход в конденсатор ТГ-3(1).txt"); STR9 = File.ReadAllLines("расход в конденсатор ТГ-3(1).txt");
            DSTR9 = new double[8]; DSTR9 = STR9.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR10 = new string[8]; streamReader = new StreamReader("расход в конденсатор ТГ-3(2).txt"); STR10 = File.ReadAllLines("расход в конденсатор ТГ-3(2).txt");
            DSTR10 = new double[8]; DSTR10 = STR10.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR11 = new string[8]; streamReader = new StreamReader("расход в конденсатор ТГ-4(1).txt"); STR11 = File.ReadAllLines("расход в конденсатор ТГ-4(1).txt");
            DSTR11 = new double[8]; DSTR11 = STR11.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR12 = new string[8]; streamReader = new StreamReader("расход в конденсатор ТГ-4(2).txt"); STR12 = File.ReadAllLines("расход в конденсатор ТГ-4(2).txt");
            DSTR12 = new double[8]; DSTR12 = STR12.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR13 = new string[8]; streamReader = new StreamReader("расход в конденсатор ТГ-6(1).txt"); STR13 = File.ReadAllLines("расход в конденсатор ТГ-6(1).txt");
            DSTR13 = new double[8]; DSTR13 = STR13.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR14 = new string[8]; streamReader = new StreamReader("расход в конденсатор ТГ-6(2).txt"); STR14 = File.ReadAllLines("расход в конденсатор ТГ-6(2).txt");
            DSTR14 = new double[8]; DSTR14 = STR14.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR15 = new string[9]; streamReader = new StreamReader("граница ТГ-1 (слева).txt"); STR15 = File.ReadAllLines("граница ТГ-1 (слева).txt");
            DSTR15 = new double[9]; DSTR15 = STR15.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR16 = new string[9]; streamReader = new StreamReader("граница ТГ-1 (справа верх).txt"); STR16 = File.ReadAllLines("граница ТГ-1 (справа верх).txt");
            DSTR16 = new double[9]; DSTR16 = STR16.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR17 = new string[9]; streamReader = new StreamReader("граница ТГ-1 (справа низ).txt"); STR17 = File.ReadAllLines("граница ТГ-1 (справа низ).txt");
            DSTR17 = new double[9]; DSTR17 = STR17.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR18 = new string[9]; streamReader = new StreamReader("граница ТГ-3 (справа 1).txt"); STR18 = File.ReadAllLines("граница ТГ-3 (справа 1).txt");
            DSTR18 = new double[9]; DSTR18 = STR18.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR19 = new string[9]; streamReader = new StreamReader("граница ТГ-3 (справа 2).txt"); STR19 = File.ReadAllLines("граница ТГ-3 (справа 2).txt");
            DSTR19 = new double[9]; DSTR19 = STR19.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR20 = new string[9]; streamReader = new StreamReader("граница ТГ-3 (слева 1).txt"); STR20 = File.ReadAllLines("граница ТГ-3 (слева 1).txt");
            DSTR20 = new double[9]; DSTR20 = STR20.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR21 = new string[9]; streamReader = new StreamReader("граница ТГ-3 (слева 2).txt"); STR21 = File.ReadAllLines("граница ТГ-3 (слева 2).txt");
            DSTR21 = new double[9]; DSTR21 = STR21.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR22 = new string[9]; streamReader = new StreamReader("граница ТГ-4 (справа 1).txt"); STR22 = File.ReadAllLines("граница ТГ-4 (справа 1).txt");
            DSTR22 = new double[9]; DSTR22 = STR22.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR23 = new string[9]; streamReader = new StreamReader("граница ТГ-4 (справа 2).txt"); STR23 = File.ReadAllLines("граница ТГ-4 (справа 2).txt");
            DSTR23 = new double[9]; DSTR23 = STR23.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR24 = new string[9]; streamReader = new StreamReader("граница ТГ-4 (слева 1).txt"); STR24 = File.ReadAllLines("граница ТГ-4 (слева 1).txt");
            DSTR24 = new double[9]; DSTR24 = STR24.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR25 = new string[9]; streamReader = new StreamReader("граница ТГ-4 (слева 2).txt"); STR25 = File.ReadAllLines("граница ТГ-4 (слева 2).txt");
            DSTR25 = new double[9]; DSTR25 = STR25.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR26 = new string[9]; streamReader = new StreamReader("граница верх ТГ-6 (N).txt"); STR26 = File.ReadAllLines("граница верх ТГ-6 (N).txt");
            DSTR26 = new double[9]; DSTR26 = STR26.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR27 = new string[4]; streamReader = new StreamReader("граница низ ТГ-6 (N).txt"); STR27 = File.ReadAllLines("граница низ ТГ-6 (N).txt");
            DSTR27 = new double[4]; DSTR27 = STR27.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR28 = new string[9]; streamReader = new StreamReader("граница справа ТГ-6 (N).txt"); STR28 = File.ReadAllLines("граница справа ТГ-6 (N).txt");
            DSTR28 = new double[9]; DSTR28 = STR28.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR29 = new string[9]; streamReader = new StreamReader("граница ТГ-4 (слева 2).txt"); STR29 = File.ReadAllLines("граница ТГ-4 (слева 2).txt");
            DSTR29 = new double[9]; DSTR29 = STR29.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR30 = new string[9]; streamReader = new StreamReader("вакуум ТГ-1 (4000).txt"); STR30 = File.ReadAllLines("вакуум ТГ-1 (4000).txt");
            DSTR30 = new double[9]; DSTR30 = STR30.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR31 = new string[9]; streamReader = new StreamReader("вакуум ТГ-1 (8000).txt"); STR31 = File.ReadAllLines("вакуум ТГ-1 (8000).txt");
            DSTR31 = new double[9]; DSTR31 = STR31.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR32 = new string[9]; streamReader = new StreamReader("вакуум ТГ-3 (4000).txt"); STR32 = File.ReadAllLines("вакуум ТГ-3 (4000).txt");
            DSTR32 = new double[9]; DSTR32 = STR32.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR33 = new string[9]; streamReader = new StreamReader("вакуум ТГ-3 (8000).txt"); STR33 = File.ReadAllLines("вакуум ТГ-3 (8000).txt");
            DSTR33 = new double[9]; DSTR33 = STR33.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR34 = new string[9]; streamReader = new StreamReader("вакуум ТГ-4 (8000).txt"); STR34 = File.ReadAllLines("вакуум ТГ-4 (8000).txt");
            DSTR34 = new double[9]; DSTR34 = STR34.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR35 = new string[9]; streamReader = new StreamReader("вакуум ТГ-4 (16000).txt"); STR35 = File.ReadAllLines("вакуум ТГ-4 (16000).txt");
            DSTR35 = new double[9]; DSTR35 = STR35.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR36 = new string[9]; streamReader = new StreamReader("вакуум ТГ-6.txt"); STR36 = File.ReadAllLines("вакуум ТГ-6.txt");
            DSTR36 = new double[9]; DSTR36 = STR36.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR37 = new string[2]; streamReader = new StreamReader("Bek14.txt"); STR37 = File.ReadAllLines("Bek14.txt");
            DSTR37 = new double[2]; DSTR37 = STR37.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR38 = new string[2]; streamReader = new StreamReader("Bek23.txt"); STR38 = File.ReadAllLines("Bek23.txt");
            DSTR38 = new double[2]; DSTR38 = STR38.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR39 = new string[2]; streamReader = new StreamReader("Bek5.txt"); STR39 = File.ReadAllLines("Bek5.txt");
            DSTR39 = new double[2]; DSTR39 = STR39.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR40 = new string[2]; streamReader = new StreamReader("Bek67.txt"); STR40 = File.ReadAllLines("Bek67.txt");
            DSTR40 = new double[2]; DSTR40 = STR40.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR41 = new string[4]; streamReader = new StreamReader("град.txt"); STR41 = File.ReadAllLines("град.txt");
            DSTR41 = new double[4]; DSTR41 = STR41.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR42 = new string[3]; streamReader = new StreamReader("граница по Qp ТГ-6 (низ).txt"); STR42 = File.ReadAllLines("граница по Qp ТГ-6 (низ).txt");
            DSTR42 = new double[3]; DSTR42 = STR42.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR43 = new string[7]; streamReader = new StreamReader("граница по Qt ТГ-6 (слева).txt"); STR43 = File.ReadAllLines("граница по Qt ТГ-6 (слева).txt");
            DSTR43 = new double[7]; DSTR43 = STR43.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR44 = new string[3]; streamReader = new StreamReader("граница по Qt ТГ-1 (справа).txt"); STR44 = File.ReadAllLines("граница по Qt ТГ-1 (справа).txt");
            DSTR44 = new double[3]; DSTR44 = STR44.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR45 = new string[2]; streamReader = new StreamReader("нагрузка ТГ-2.txt"); STR45 = File.ReadAllLines("нагрузка ТГ-2.txt");
            DSTR45 = new double[2]; DSTR45 = STR45.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR46 = new string[2]; streamReader = new StreamReader("нагрузка ТГ-5.txt"); STR46 = File.ReadAllLines("нагрузка ТГ-5.txt");
            DSTR46 = new double[2]; DSTR46 = STR46.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR47 = new string[2]; streamReader = new StreamReader("ПТВМ.txt"); STR47 = File.ReadAllLines("ПТВМ.txt");
            DSTR47 = new double[2]; DSTR47 = STR47.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR48 = new string[3]; streamReader = new StreamReader("СНПТВМ.txt"); STR48 = File.ReadAllLines("СНПТВМ.txt");
            DSTR48 = new double[3]; DSTR48 = STR48.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR49 = new string[2]; streamReader = new StreamReader("ТГ-2(е).txt"); STR49 = File.ReadAllLines("ТГ-2(е).txt");
            DSTR49 = new double[2]; DSTR49 = STR49.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR50 = new string[2]; streamReader = new StreamReader("ТГ-5(е).txt"); STR50 = File.ReadAllLines("ТГ-5(е).txt");
            DSTR50 = new double[2]; DSTR50 = STR50.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR51 = new string[27]; streamReader = new StreamReader("ТГ-6(Q(D)).txt"); STR51 = File.ReadAllLines("ТГ-6(Q(D)).txt");
            DSTR51 = new double[27]; DSTR51 = STR51.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR52 = new string[2]; streamReader = new StreamReader("ТГ-2.txt"); STR52 = File.ReadAllLines("ТГ-2.txt");
            DSTR52 = new double[2]; DSTR52 = STR52.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR53 = new string[2]; streamReader = new StreamReader("ТГ-5.txt"); STR53 = File.ReadAllLines("ТГ-5.txt");
            DSTR53 = new double[2]; DSTR53 = STR53.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR54 = new string[2]; streamReader = new StreamReader("ТГ-6(Q).txt"); STR54 = File.ReadAllLines("ТГ-6(Q).txt");
            DSTR54 = new double[2]; DSTR54 = STR54.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR56 = new string[2]; streamReader = new StreamReader("ТГ-1 (поправка на температуру).txt"); STR56 = File.ReadAllLines("ТГ-1 (поправка на температуру).txt");
            DSTR56 = new double[2]; DSTR56 = STR56.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR60 = new string[9]; streamReader = new StreamReader("граница ТГ-1 по Qt (слева).txt"); STR60 = File.ReadAllLines("граница ТГ-1 по Qt (слева).txt");
            DSTR60 = new double[9]; DSTR60 = STR60.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR61 = new string[9]; streamReader = new StreamReader("граница ТГ-1 по Qt (справа верх).txt"); STR61 = File.ReadAllLines("граница ТГ-1 по Qt (справа верх).txt");
            DSTR61 = new double[9]; DSTR61 = STR61.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR62 = new string[9]; streamReader = new StreamReader("граница ТГ-1 по Qt (справа низ).txt"); STR62 = File.ReadAllLines("граница ТГ-1 по Qt (справа низ).txt");
            DSTR62 = new double[9]; DSTR62 = STR62.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR63 = new string[9]; streamReader = new StreamReader("граница ТГ-3 по Qt (слева верх).txt"); STR63 = File.ReadAllLines("граница ТГ-3 по Qt (слева верх).txt");
            DSTR63 = new double[9]; DSTR63 = STR63.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR64 = new string[9]; streamReader = new StreamReader("граница ТГ-3 по Qt (слева низ).txt"); STR64 = File.ReadAllLines("граница ТГ-3 по Qt (слева низ).txt");
            DSTR64 = new double[9]; DSTR64 = STR64.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR65 = new string[9]; streamReader = new StreamReader("граница ТГ-3 по Qt (справа верх).txt"); STR65 = File.ReadAllLines("граница ТГ-3 по Qt (справа верх).txt");
            DSTR65 = new double[9]; DSTR65 = STR65.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR66 = new string[9]; streamReader = new StreamReader("граница ТГ-3 по Qt (справа низ).txt"); STR66 = File.ReadAllLines("граница ТГ-3 по Qt (справа низ).txt");
            DSTR66 = new double[9]; DSTR66 = STR66.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR67 = new string[9]; streamReader = new StreamReader("граница ТГ-4 по Qt (слева 1).txt"); STR67 = File.ReadAllLines("граница ТГ-4 по Qt (слева 1).txt");
            DSTR67 = new double[9]; DSTR67 = STR67.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR68 = new string[9]; streamReader = new StreamReader("граница ТГ-4 по Qt (слева 2).txt"); STR68 = File.ReadAllLines("граница ТГ-4 по Qt (слева 2).txt");
            DSTR68 = new double[9]; DSTR68 = STR68.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR69 = new string[9]; streamReader = new StreamReader("граница ТГ-4 по Qt (справа 1).txt"); STR69 = File.ReadAllLines("граница ТГ-4 по Qt (справа 1).txt");
            DSTR69 = new double[9]; DSTR69 = STR69.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR70 = new string[9]; streamReader = new StreamReader("граница ТГ-4 по Qt (справа 2).txt"); STR70 = File.ReadAllLines("граница ТГ-4 по Qt (справа 2).txt");
            DSTR70 = new double[9]; DSTR70 = STR70.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
            STR71 = new string[3]; streamReader = new StreamReader("граница по N ТГ-6 (слева).txt"); STR71 = File.ReadAllLines("граница по N ТГ-6 (слева).txt");
            DSTR71 = new double[3]; DSTR71 = STR71.Select(x => double.Parse(x, CultureInfo.GetCultureInfo("en-US"))).ToArray();
        }

        public void Null()
        {
            h30n = 0; N1 = 0; Qt1 = 0; D1t = 0; Dp2 = 0; N3 = 0; Qt3 = 0; N4 = 0; Qt4 = 0; Dp5 = 0; N6 = 0; Qt6 = 0; Dp6 = 0; D140_13 = 0; D140_30 = 0; Dp30_13 = 0; D3013 = 0; Qpb = 0; Qpvk = 0; Dptn1 = 0; Dptn2 = 0; lt = 0; lp = 0; ln = 0; lp30 = 0; N5 = 0; N2 = 0; Dv6 = 0; Dv2 = 0; Dv51 = 0; Dv52 = 0; D30_13 = 0; Dv30_13 = 0; Dv140_13 = 0; Qpsn = 0;
            Qs = 0; P6 = 0; P4 = 0; P3 = 0; P1 = 0; tcw1 = 0; tcw4 = 0; tcw3 = 0; tcw6 = 0; tcwo3 = 0; tcwo4 = 0; tcwo6 = 0; k1 = 0; k2 = 0; k3 = 0; k4 = 0; k5 = 0; k6 = 0; k7 = 0; Gek = 0; G1p = 0; G2p = 0; G3p = 0; G4p = 0; G5p = 0; G6p = 0; G7p = 0; nk = 0; S = 0; Nko = 0; Ncn = 0; Nto = 0; anas = 0; apvk = 0; Qe = 0; Qsnto = 0; Ke = 0;
            Nsnee = 0; Nsn = 0; B = 0; Be = 0; ntp = 0; Kotrk = 0; nnet = 0; be = 0; bte = 0; dQe = 0; Kotrt = 0; Qsnko = 0; Dptn1 = 0; Dptn2 = 0; Ddvd = 0; Ddnd = 0; Dpsv = 0; lN31 = 0; lN32 = 0; lN11 = 0; lN12 = 0; lN13 = 0; lN41 = 0; lN42 = 0; lN43 = 0; lN44 = 0; lN45 = 0; lN61 = 0; lN62 = 0; lN63 = 0; Q1t = 0; MM = 0; MMM = 0; MMM1 = 0; MMM2 = 0;
            a = 0; dptn1 = 0; dptn2 = 0; tp1 = 0; tp3 = 0; tp4 = 0; tp6 = 0; np1 = 0; np3 = 0; np4 = 0; np6 = 0; dp2 = 0; dp5 = 0; dp6 = 0; dp3013 = 0; Fs = 0; D = 0; D11 = 0; rty = 0; q = 0; dtk1 = 0; dtk3 = 0; dtk4 = 0; dtk6 = 0; dtsp1 = 0; dtsp3 = 0; dtsp4 = 0; dtsp6 = 0; ogrN1 = 0; ogrN3 = 0; ogrN4 = 0; ogrN6 = 0; ogrQt1 = 0;
            ogrQt3 = 0; ogrQt4 = 0; ogrQt6 = 0; ch1 = 0; ch2 = 0; ch3 = 0; ch4 = 0; ch5 = 0; indcn1 = 0; indcn3 = 0; indcn4 = 0; Nv1 = 0; Nv3 = 0; Nv4 = 0; Nv6 = 0; if (checkBox10.IsChecked == false) Ns = 0; rtn = 0; in6 = 0; indG1 = 0; indG3 = 0; indG4 = 0; indG6 = 0;
            Array.Clear(X, 0, 35); Array.Clear(Y, 0, 42); Array.Clear(F, 0, 24); Array.Clear(X2, 0, 100); Array.Clear(F1, 0, 24); Array.Clear(Pr, 0, 24); Array.Clear(VAR, 0, 100);
            if (combobox9.Text == "активен" | combobox10.Text == "активен")
            {
                radioButton1.IsChecked = false; radioButton2.IsChecked = false; radioButton8.IsChecked = false;
                radioButton10.IsChecked = false; radioButton12.IsChecked = false; radioButton13.IsChecked = false;
            }
        }

        public int ogrTG6var(int nv)
        {
            int nn = 0;
            if (Dp6 == 0 & Qt6 == 0 & N6 == 0) return nn;
            else
            {
                if (checkBox10.IsChecked == true & (nv == 2 || nv == 6)) { nn = 1; return nn; }
                if (Dp6 > Dp6max) Dp6 = Dp6max;
                if (nv == 1)
                {
                    if (Dp6 > ogrTG6db(N6))
                    {
                        if (dp6 != 4) Dp6 = ogrTG6db(N6);
                        else { nn = 1; return nn; }
                    }
                    if (Dp6 < ogrTG6dn(N6)) { nn = 1; return nn; }
                    if (Dp6 > ogr_TG6D(N6))
                    {
                        if (dp6 != 4) Dp6 = ogr_TG6D(N6);
                        else { nn = 1; return nn; }
                    }
                    if (tp6 != 4)
                    {
                        Nv6 = ogrTG6n(Dp6);
                        if (N6 < Nv6 & Qt6 > ogr_TG6Dp(Dp6) & Qt6 > ogr_TG6Qtln(Dp6, N6) + 1) Qt6 = ogr_TG6Qtln(Dp6, N6);
                        if (N6 < Nv6 & Qt6 < ogr_TG6Dp(Dp6) & Qt6 < ogr_TG6Qtlb(Dp6, N6) - 1) Qt6 = ogr_TG6Qtlb(Dp6, N6);
                        if (N6 > Nv6 & Qt6 > ogr_TG6Qtr(Dp6, N6) + 1) Qt6 = ogr_TG6Qtr(Dp6, N6);
                        if (Qt6 < 0) Qt6 = 0;
                        Qpb = Qt - Qpodp - Qt1 - Qt3 - Qt4 - Qt6;
                    }
                    else
                    {
                        if (Qt6 > ogr_TG6Qtr(263, N6) + 1) { nn = 1; return nn; }
                        if (Qt6 > ogr_TG6Qtln(0, N6) + 1) { nn = 1; return nn; }
                        if (Qt6 < ogr_TG6Qtlb(0, N6) - 1) { nn = 1; return nn; }
                        if (ogr_TG6Qprb(Dp6, Qt6, N6))
                        {
                            if (dp6 != 4) { Dp6 = ogr_TG6Qpr(Dp6, Qt6, N6); if (Dp6 == 1001) { nn = 1; return nn; } }
                            else { nn = 1; return nn; }
                        }
                        if (Qt6 > ogr_TG6Dp(Dp6))
                        {
                            if (Dp6 > ogr_TG6Qpln(Qt6, N6) & dp6 != 4) Dp6 = ogr_TG6Qpln(Qt6, N6);
                            else { nn = 1; return nn; }
                        }
                        else
                        {
                            if (Dp6 > ogr_TG6Qplb(Qt6, N6) & dp6 != 4) Dp6 = ogr_TG6Qplb(Qt6, N6);
                            else { nn = 1; return nn; }
                        }
                    }
                }
                if (nv == 2)
                {
                    if (Qt6 > ogr_TG6bd(Dp6))
                    {
                        if (tp6 != 4) Qt6 = ogr_TG6bd(Dp6);
                        else { nn = 1; return nn; }
                        Qpb = Qt - Qpodp - Qt1 - Qt3 - Qt4 - Qt6;
                    }
                    if (np6 != 4)
                    {
                        if (N6 > ogr_TG6r(Dp6, Qt6)) N6 = ogr_TG6r(Dp6, Qt6);
                        if (N6 < ogr_TG6ln(Dp6, Qt6)) N6 = ogr_TG6ln(Dp6, Qt6);
                        if (N6 < ogr_TG6lb(Dp6, Qt6)) N6 = ogr_TG6lb(Dp6, Qt6);
                    }
                    else
                    {
                        if (N6 > ogr_TG6r(Dp6, 0)) { nn = 1; return nn; }
                        if (N6 < ogr_TG6lb(Dp6, ogr_TG6Dp(Dp6))) { nn = 1; return nn; }
                        Nv6 = ogrTG6n(Dp6);
                        if (N6 < Nv6 & Qt6 > ogr_TG6Dp(Dp6) & Qt6 > ogr_TG6Qtln(Dp6, N6) + 1)
                        {
                            if (tp6 != 4) Qt6 = ogr_TG6Qtln(Dp6, N6);
                            else { nn = 1; return nn; }
                            Qpb = Qt - Qpodp - Qt1 - Qt3 - Qt4 - Qt6;
                        }
                        if (N6 < Nv6 & Qt6 < ogr_TG6Dp(Dp6) & Qt6 < ogr_TG6Qtlb(Dp6, N6) - 1)
                        {
                            if (tp6 != 4) Qt6 = ogr_TG6Qtlb(Dp6, N6);
                            else { nn = 1; return nn; }
                            Qpb = Qt - Qpodp - Qt1 - Qt3 - Qt4 - Qt6;
                        }
                        if (N6 > Nv6 & Qt6 > ogr_TG6Qtr(Dp6, N6) + 1)
                        {
                            if (tp6 != 4) Qt6 = ogr_TG6Qtr(Dp6, N6);
                            else { nn = 1; return nn; }
                            if (Qt6 < 0) Qt6 = 0;
                            Qpb = Qt - Qpodp - Qt1 - Qt3 - Qt4 - Qt6;
                        }
                    }
                }
                if (nv == 3)
                {
                    if (N6 > ogr_TG6r(263, Qt6))
                    {
                        if (np6 != 4) N6 = ogr_TG6r(263, Qt6);
                        else { nn = 1; return nn; }
                    }
                    if (N6 < ogr_TG6lb(0, Qt6))
                    {
                        if (np6 != 4) N6 = ogr_TG6lb(0, Qt6);
                        else { nn = 1; return nn; }
                    }
                    if (N6 < ogr_TG6ln(0, Qt6))
                    {
                        if (np6 != 4) N6 = ogr_TG6ln(0, Qt6);
                        else { nn = 1; return nn; }
                    }
                    if (dp6 != 4)
                    {
                        if (ogr_TG6Qprb(Dp6, Qt6, N6))
                        {
                            if (dp6 != 4) { Dp6 = ogr_TG6Qpr(Dp6, Qt6, N6); if (Dp6 == 1001) { nn = 1; return nn; } }
                            else { nn = 1; return nn; }
                        }
                        if (Qt6 > ogr_TG6Dp(Dp6))
                        {
                            if (Dp6 > ogr_TG6Qpln(Qt6, N6) & dp6 != 4) Dp6 = ogr_TG6Qpln(Qt6, N6);
                            else { nn = 1; return nn; }
                        }
                        else
                        {
                            if (Dp6 > ogr_TG6Qplb(Qt6, N6) & dp6 != 4) Dp6 = ogr_TG6Qplb(Qt6, N6);
                            else { nn = 1; return nn; }
                        }
                    }
                    else
                    {
                        if (Dp6 > ogr_TG6Qt(Qt6)) { nn = 1; return nn; }
                        if (N6 > ogr_TG6r(Dp6, Qt6))
                        {
                            if (np6 != 4) N6 = ogr_TG6r(Dp6, Qt6);
                            else { nn = 1; return nn; }
                        }
                        if (N6 < ogr_TG6lb(Dp6, Qt6))
                        {
                            if (np6 != 4) N6 = ogr_TG6lb(Dp6, Qt6);
                            else { nn = 1; return nn; }
                        }

                        if (N6 < ogr_TG6ln(Dp6, Qt6))
                        {
                            if (np6 != 4) N6 = ogr_TG6ln(Dp6, Qt6);
                            else { nn = 1; return nn; }
                        }
                    }
                }
                if (nv == 4)
                {
                    if (Qt6 > ogr_TG6Qtr(263, N6) + 1)
                    {
                        if (tp6 != 4) Qt6 = ogr_TG6Qtr(263, N6);
                        else { nn = 1; return nn; }
                        if (Qt6 < 0) Qt6 = 0;
                        Qpb = Qt - Qpodp - Qt1 - Qt3 - Qt4 - Qt6;
                    }
                    if (Qt6 > ogr_TG6Qtln(0, N6) + 1)
                    {
                        if (tp6 != 4) Qt6 = ogr_TG6Qtln(0, N6);
                        else { nn = 1; return nn; }
                        Qpb = Qt - Qpodp - Qt1 - Qt3 - Qt4 - Qt6;
                    }
                    if (Qt6 < ogr_TG6Qtlb(0, N6))
                    {
                        if (tp6 != 4) Qt6 = ogr_TG6Qtlb(0, N6);
                        else { nn = 1; return nn; }
                        Qpb = Qt - Qpodp - Qt1 - Qt3 - Qt4 - Qt6;
                    }
                    if (dp6 != 4)
                    {
                        if (ogr_TG6Qprb(Dp6, Qt6, N6))
                        {
                            if (dp6 != 4) { Dp6 = ogr_TG6Qpr(Dp6, Qt6, N6); if (Dp6 == 1001) { nn = 1; return nn; } }
                            else { nn = 1; return nn; }
                        }
                        if (Qt6 > ogr_TG6Dp(Dp6))
                        {
                            if (Dp6 > ogr_TG6Qpln(Qt6, N6) & dp6 != 4) Dp6 = ogr_TG6Qpln(Qt6, N6);
                            else { nn = 1; return nn; }
                        }
                        else
                        {
                            if (Dp6 > ogr_TG6Qplb(Qt6, N6) & dp6 != 4) Dp6 = ogr_TG6Qplb(Qt6, N6);
                            else { nn = 1; return nn; }
                        }
                    }
                    else
                    {
                        if (Dp6 < ogrTG6dn(N6) || Dp6 > ogrTG6db(N6)) { nn = 1; return nn; }
                        Nv6 = ogrTG6n(Dp6);
                        if (N6 < Nv6 & Qt6 > ogr_TG6Dp(Dp6) & Qt6 > ogr_TG6Qtln(Dp6, N6) + 1)
                        {
                            if (tp6 != 4) Qt6 = ogr_TG6Qtln(Dp6, N6);
                            else { nn = 1; return nn; }
                            Qpb = Qt - Qpodp - Qt1 - Qt3 - Qt4 - Qt6;
                        }
                        if (N6 < Nv6 & Qt6 < ogr_TG6Dp(Dp6) & Qt6 < ogr_TG6Qtlb(Dp6, N6) - 1)
                        {
                            if (tp6 != 4) Qt6 = ogr_TG6Qtlb(Dp6, N6);
                            else { nn = 1; return nn; }
                            Qpb = Qt - Qpodp - Qt1 - Qt3 - Qt4 - Qt6;
                        }
                        if (N6 > Nv6 & Qt6 > ogr_TG6Qtr(Dp6, N6) + 1)
                        {
                            if (tp6 != 4) Qt6 = ogr_TG6Qtr(Dp6, N6);
                            else { nn = 1; return nn; }
                            if (Qt6 < 0) Qt6 = 0;
                            Qpb = Qt - Qpodp - Qt1 - Qt3 - Qt4 - Qt6;
                        }
                    }
                }
                if (nv == 5)
                {
                    if (N6 > ogr_TG6r(Dp6, 0))
                    {
                        if (np6 != 4) N6 = ogr_TG6r(Dp6, 0);
                        else { nn = 1; return nn; }
                    }
                    if (N6 < ogr_TG6lb(Dp6, ogr_TG6Dp(Dp6)))
                    {
                        if (np6 != 4) N6 = ogr_TG6lb(Dp6, ogr_TG6Dp(Dp6));
                        else { nn = 1; return nn; }
                    }
                    if (tp6 != 4)
                    {
                        Nv6 = ogrTG6n(Dp6);
                        if (N6 < Nv6 & Qt6 > ogr_TG6Dp(Dp6) & Qt6 > ogr_TG6Qtln(Dp6, N6) + 1) Qt6 = ogr_TG6Qtln(Dp6, N6);
                        if (N6 < Nv6 & Qt6 < ogr_TG6Dp(Dp6) & Qt6 < ogr_TG6Qtlb(Dp6, N6) - 1) Qt6 = ogr_TG6Qtlb(Dp6, N6);
                        if (N6 > Nv6 & Qt6 > ogr_TG6Qtr(Dp6, N6) + 1) Qt6 = ogr_TG6Qtr(Dp6, N6);
                        if (Qt6 < 0) Qt6 = 0;
                        Qpb = Qt - Qpodp - Qt1 - Qt3 - Qt4 - Qt6;
                    }
                    else
                    {
                        if (Qt6 > ogr_TG6Dp(Dp6)) { nn = 1; return nn; }
                        if (N6 > ogr_TG6r(Dp6, Qt6))
                        {
                            if (np6 != 4) N6 = ogr_TG6r(Dp6, Qt6);
                            else { nn = 1; return nn; }
                        }
                        if (N6 < ogr_TG6lb(Dp6, Qt6))
                        {
                            if (np6 != 4) N6 = ogr_TG6lb(Dp6, Qt6);
                            else { nn = 1; return nn; }
                        }
                        if (N6 < ogr_TG6ln(Dp6, Qt6))
                        {
                            if (np6 != 4) N6 = ogr_TG6ln(Dp6, Qt6);
                            else { nn = 1; return nn; }
                        }
                    }
                }
                if (nv == 6)
                {
                    if (Dp6 > ogr_TG6bt(Qt6))
                    {
                        if (dp6 != 4) Dp6 = ogr_TG6bt(Qt6);
                        else { nn = 1; return nn; }
                    }
                    if (np6 != 4)
                    {
                        if (N6 > ogr_TG6r(Dp6, Qt6)) N6 = ogr_TG6r(Dp6, Qt6);
                        if (N6 < ogr_TG6lb(Dp6, Qt6)) N6 = ogr_TG6lb(Dp6, Qt6);
                        if (N6 < ogr_TG6ln(Dp6, Qt6)) N6 = ogr_TG6ln(Dp6, Qt6);
                    }
                    else
                    {
                        if (N6 > ogr_TG6r(263, Qt6)) { nn = 1; return nn; }
                        if (Qt6 < ogr_TG6Dp(0) && N6 < ogr_TG6lb(0, Qt6)) { nn = 1; return nn; }
                        if (Qt6 > ogr_TG6Dp(0) && N6 < ogr_TG6ln(0, Qt6)) { nn = 1; return nn; }
                        if (ogr_TG6Qprb(Dp6, Qt6, N6))
                        {
                            if (dp6 != 4) { Dp6 = ogr_TG6Qpr(Dp6, Qt6, N6); if (Dp6 == 1001) { nn = 1; return nn; } }
                            else { nn = 1; return nn; }
                        }
                        if (Qt6 > ogr_TG6Dp(Dp6))
                        {
                            if (Dp6 > ogr_TG6Qpln(Qt6, N6) & dp6 != 4) Dp6 = ogr_TG6Qpln(Qt6, N6);
                            else { nn = 1; return nn; }
                        }
                        else
                        {
                            if (Dp6 > ogr_TG6Qplb(Qt6, N6) & dp6 != 4) Dp6 = ogr_TG6Qplb(Qt6, N6);
                            else { nn = 1; return nn; }
                        }
                    }
                }
            }
            return nn;
        }

        public void Replace()
    {
        textBox1.Text = textBox1.Text.Replace('.', ','); textBox2.Text = textBox2.Text.Replace('.', ','); textBox3.Text = textBox3.Text.Replace('.', ','); textBox4.Text = textBox4.Text.Replace('.', ',');
        textBox5.Text = textBox5.Text.Replace('.', ','); textBox6.Text = textBox6.Text.Replace('.', ','); textBox7.Text = textBox7.Text.Replace('.', ','); textBox8.Text = textBox8.Text.Replace('.', ',');
        textBox9.Text = textBox9.Text.Replace('.', ','); textBox10.Text = textBox10.Text.Replace('.', ','); textBox11.Text = textBox11.Text.Replace('.', ','); textBox12.Text = textBox12.Text.Replace('.', ',');
        textBox14.Text = textBox14.Text.Replace('.', ','); textBox28.Text = textBox28.Text.Replace('.', ','); textBox29.Text = textBox29.Text.Replace('.', ','); textBox30.Text = textBox30.Text.Replace('.', ',');
        textBox70.Text = textBox70.Text.Replace('.', ','); textBox22.Text = textBox22.Text.Replace('.', ','); textBox18.Text = textBox18.Text.Replace('.', ','); textBox19.Text = textBox19.Text.Replace('.', ',');
        textBox21.Text = textBox21.Text.Replace('.', ','); textBox20.Text = textBox20.Text.Replace('.', ','); textBox13.Text = textBox13.Text.Replace('.', ','); textBox39.Text = textBox39.Text.Replace('.', ',');
        textBox57.Text = textBox57.Text.Replace('.', ','); textBox58.Text = textBox58.Text.Replace('.', ','); textBox59.Text = textBox59.Text.Replace('.', ','); textBox60.Text = textBox60.Text.Replace('.', ','); textBox61.Text = textBox61.Text.Replace('.', ',');
    }
                
        private void Button2_Click(object sender, RoutedEventArgs e)
        {
            if (combobox9.Text == "активен")
            {
                excelapp.Visible = true;
                excelappworkbook = excelapp.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + @"\Макет.xlsx",
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
        }
            
    }

    public class MatrixDecompositionProgram
    {
        public static double[][] MatrixCreate(int rows, int cols)
        {
            double[][] result = new double[rows][];
            for (int i = 0; i < rows; ++i)
                result[i] = new double[cols];
            return result;
        }
        public static double[][] MatrixDecompose(double[][] matrix, out int[] perm, out int toggle)
        {
            int rows = 24;
            int cols = 24;
            if (rows != cols)
                throw new Exception("Attempt to MatrixDecompose a non-square mattrix");

            int n = rows;

            double[][] result = MatrixDuplicate(matrix);
            perm = new int[n];
            for (int i = 0; i < n; ++i) { perm[i] = i; }
            toggle = 1;

            for (int j = 0; j < n - 1; ++j) // each column
            {
                double colMax = Math.Abs(result[j][j]);
                int pRow = j;
                for (int i = j + 1; i < n; ++i)
                {
                    if (result[i][j] > colMax)
                    {
                        colMax = result[i][j];
                        pRow = i;
                    }
                }

                if (pRow != j)
                {
                    double[] rowPtr = result[pRow];
                    result[pRow] = result[j];
                    result[j] = rowPtr;

                    int tmp = perm[pRow]; // and swap perm info
                    perm[pRow] = perm[j];
                    perm[j] = tmp;

                    toggle = -toggle; // adjust the row-swap toggle
                }

                if (Math.Abs(result[j][j]) < 1.0E-20) // if diagonal after swap is zero . . .
                    return null; // consider a throw

                for (int i = j + 1; i < n; ++i)
                {
                    result[i][j] /= result[j][j];
                    for (int k = j + 1; k < n; ++k)
                    {
                        result[i][k] -= result[i][j] * result[j][k];
                    }
                }
            } // main j column loop

            return result;
        }
        public static double[][] MatrixInverse(double[][] matrix)
        {
            int n = matrix.Length;
            double[][] result = MatrixDuplicate(matrix);

            int[] perm;
            int toggle;
            double[][] lum = MatrixDecompose(matrix, out perm, out toggle);
            if (lum == null)
                throw new Exception("Unable to compute inverse");

            double[] b = new double[n];
            for (int i = 0; i < n; ++i)
            {
                for (int j = 0; j < n; ++j)
                {
                    if (i == perm[j])
                        b[j] = 1.0;
                    else
                        b[j] = 0.0;
                }

                double[] x = HelperSolve(lum, b);

                for (int j = 0; j < n; ++j)
                    result[j][i] = x[j];
            }
            return result;
        }
        public static double MatrixDeterminant(double[][] matrix)
        {
            int[] perm;
            int toggle;
            double[][] lum = MatrixDecompose(matrix, out perm, out toggle);
            if (lum == null)
                throw new Exception("Unable to compute MatrixDeterminant");
            double result = toggle;
            for (int i = 0; i < lum.Length; ++i)
                result *= lum[i][i];
            return result;
        }
        public static double[] HelperSolve(double[][] luMatrix, double[] b)
        {
            int n = luMatrix.Length;
            double[] x = new double[n];
            b.CopyTo(x, 0);

            for (int i = 1; i < n; ++i)
            {
                double sum = x[i];
                for (int j = 0; j < i; ++j)
                    sum -= luMatrix[i][j] * x[j];
                x[i] = sum;
            }

            x[n - 1] /= luMatrix[n - 1][n - 1];
            for (int i = n - 2; i >= 0; --i)
            {
                double sum = x[i];
                for (int j = i + 1; j < n; ++j)
                    sum -= luMatrix[i][j] * x[j];
                x[i] = sum / luMatrix[i][i];
            }

            return x;
        }
        public static double[][] MatrixDuplicate(double[][] matrix)
        {
            double[][] result = MatrixCreate(matrix.Length, matrix[0].Length);
            for (int i = 0; i < matrix.Length; ++i)
                for (int j = 0; j < matrix[i].Length; ++j)
                    result[i][j] = matrix[i][j];
            return result;
        }
        public static double[][] ExtractLower(double[][] matrix)
        {
            int rows = matrix.Length; int cols = matrix[0].Length;
            double[][] result = MatrixCreate(rows, cols);
            for (int i = 0; i < rows; ++i)
            {
                for (int j = 0; j < cols; ++j)
                {
                    if (i == j)
                        result[i][j] = 1.0;
                    else if (i > j)
                        result[i][j] = matrix[i][j];
                }
            }
            return result;
        }
        public static double[][] ExtractUpper(double[][] matrix)
        {
            int rows = matrix.Length; int cols = matrix[0].Length;
            double[][] result = MatrixCreate(rows, cols);
            for (int i = 0; i < rows; ++i)
            {
                for (int j = 0; j < cols; ++j)
                {
                    if (i <= j)
                        result[i][j] = matrix[i][j];
                }
            }
            return result;
        }
    }
}
