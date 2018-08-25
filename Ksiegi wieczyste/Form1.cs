using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using AutoIt;
using AutoItX3Lib;
using System.Data.SqlClient;
using HtmlAgilityPack;

namespace Ksiegi_wieczyste
{
    public partial class Form1 : Form
    {
        private string _dzial_III = null;
        private string _dzial_II = null;
        private string _dzial_Isp = null;
        private string _dzial_Io = null;

       

        public string dzial_III
        {
            get { return _dzial_III; }
            set { _dzial_III = value; }
        }

        public string dzial_II
        {
            get { return _dzial_II; }
            set { _dzial_II = value; }
        }

        public string dzial_Isp
        {
            get { return _dzial_Isp; }
            set { _dzial_Isp = value; }
        }

        public string dzial_Io
        {
            get { return _dzial_Io; }
            set { _dzial_Io = value; }
        }



        public Form1()
        {
            InitializeComponent();
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            

            if(AutoItX.WinActivate("EUKW - Prezentacja Księgi Wieczystej - Mozilla Firefox", "") == 1)
            {
                MessageBox.Show("okno aktywne");
            } else
            {
                DialogResult dialogResult = MessageBox.Show("Czy uruchomić stronę EKW?", "Komunikat", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    var iPID = AutoItX.Run(@"C:\Program Files (x86)\Mozilla Firefox\firefox.exe https://przegladarka-ekw.ms.gov.pl/eukw_prz/KsiegiWieczyste/wyszukiwanieKW?komunikaty=true&kontakt=true&okienkoSerwisowe=false", "", 1);
                }
            }
        }


        private void wyszukajKsiege(string _kw)
        {
            string[] separator = {@"/" };
            string[] result = _kw.Split(separator, StringSplitOptions.RemoveEmptyEntries);
            try
            {
                if (AutoItX.WinActivate("EUKW - Prezentacja Księgi Wieczystej - Mozilla Firefox", "") == 1)
                {
                    string numer_ksiegi = result[1];
                    string cyfra_kontrolna = result[2];
                    var notification = new System.Windows.Forms.NotifyIcon()
                    {
                        Visible = true,
                        Icon = System.Drawing.SystemIcons.Information,
                        BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info,
                        BalloonTipTitle = "AutoIT - Ksiei Wieczyste",
                        BalloonTipText = "Trwa pobieranie danych z księgi wieczystej: SR2W/" + numer_ksiegi + '/' + cyfra_kontrolna,
                    };
                    notification.ShowBalloonTip(5000);

                    //Przesuń i wpisz sąd
                    AutoItX.MouseMove(2433, 464, 1);
                    AutoItX.Sleep(500);
                    AutoItX.MouseClick();
                    AutoItX.Send("SR2W");
                    AutoItX.Send("{TAB}");
                    AutoItX.Sleep(500);

                    //Wpisz numer księgi            
                    AutoItX.Send(numer_ksiegi);
                    AutoItX.Send("{TAB}");
                    AutoItX.Sleep(500);

                    //Wpisz cyfrę kontrolną
                    AutoItX.Send(cyfra_kontrolna);
                    AutoItX.Sleep(500);

                    //przesuń nad google captcha
                    AutoItX.MouseMove(2433, 546);
                    AutoItX.Sleep(500);
                    AutoItX.MouseClick();
                    AutoItX.Sleep(4000);

                    //wyszukaj księge
                    AutoItX.MouseMove(2933, 685);
                    AutoItX.Sleep(500);
                    AutoItX.MouseClick();
                    AutoItX.Sleep(6000);

                    //Przeglądanie aktualnej treści KW
                    AutoItX.MouseMove(2229, 797);
                    AutoItX.Sleep(500);
                    AutoItX.MouseClick();
                    AutoItX.Sleep(5000);

                    //Dzial III
                    AutoItX.MouseMove(2512, 219);
                    AutoItX.MouseClick();
                    AutoItX.Sleep(600);
                    AutoItX.Send("^u");
                    AutoItX.Sleep(600);
                    AutoItX.Send("^a");
                    AutoItX.Sleep(500);
                    AutoItX.Send("^c");
                    AutoItX.Send("^{F4}");
                    AutoItX.Sleep(1000);
                    dzial_III = Clipboard.GetText();
                    AutoItX.Sleep(500);

                    //DZIAL II
                    AutoItX.MouseMove(2349, 219);
                    AutoItX.Sleep(500);
                    AutoItX.MouseClick();
                    AutoItX.Sleep(500);
                    AutoItX.Send("^u");
                    AutoItX.Sleep(600);
                    AutoItX.Send("^a");
                    AutoItX.Sleep(500);
                    AutoItX.Send("^c");
                    AutoItX.Send("^{F4}");
                    AutoItX.Sleep(1000);
                    dzial_II = Clipboard.GetText();
                    AutoItX.Sleep(500);

                    //DZIAL I-sp
                    AutoItX.MouseMove(2176, 219);
                    AutoItX.Sleep(500);
                    AutoItX.MouseClick();
                    AutoItX.Sleep(500);
                    AutoItX.Send("^u");
                    AutoItX.Sleep(600);
                    AutoItX.Send("^a");
                    AutoItX.Sleep(500);
                    AutoItX.Send("^c");
                    AutoItX.Send("^{F4}");
                    AutoItX.Sleep(1000);
                    dzial_Isp = Clipboard.GetText();
                    AutoItX.Sleep(500);

                    //DZIAL I-O
                    AutoItX.MouseMove(1970, 219);
                    AutoItX.Sleep(500);
                    AutoItX.MouseClick();
                    AutoItX.Sleep(500);
                    AutoItX.Send("^u");
                    AutoItX.Sleep(600);
                    AutoItX.Send("^a");
                    AutoItX.Sleep(500);
                    AutoItX.Send("^c");
                    AutoItX.Send("^{F4}");
                    AutoItX.Sleep(1000);
                    dzial_Io = Clipboard.GetText();
                    AutoItX.Sleep(500);


                    //Powrót do początku

                    AutoItX.Send("{END}");
                    AutoItX.MouseMove(1970, 946);
                    AutoItX.Sleep(500);
                    AutoItX.MouseClick();
                    AutoItX.Sleep(1500);
                    AutoItX.Send("{END}");
                    AutoItX.Sleep(500);
                    AutoItX.MouseMove(2662, 748);
                    AutoItX.Sleep(500);
                    AutoItX.MouseClick();

                    richTextBox1.AppendText(dzial_III);
                    richTextBox2.AppendText(dzial_II);
                    richTextBox3.AppendText(dzial_Isp);
                    richTextBox4.AppendText(dzial_Io);

                    StringBuilder numerKsiegi = new StringBuilder();
                    numerKsiegi.Append("SR2W/");
                    numerKsiegi.Append(numer_ksiegi);
                    numerKsiegi.Append("/");
                    numerKsiegi.Append(cyfra_kontrolna);

                    dodajKsiegeDoBazy(numerKsiegi.ToString(), dzial_III, dzial_II, dzial_Isp, dzial_Io);

                }
                else
                {
                    DialogResult dialogResult = MessageBox.Show("Czy uruchomić stronę EKW?", "Komunikat", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        var iPID = AutoItX.Run(@"C:\Program Files (x86)\Mozilla Firefox\firefox.exe https://przegladarka-ekw.ms.gov.pl/eukw_prz/KsiegiWieczyste/wyszukiwanieKW?komunikaty=true&kontakt=true&okienkoSerwisowe=false", "", 1);
                    }
                }
            }
            catch (Exception x)
            {
                MessageBox.Show(x.ToString());
                
                throw;
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (AutoItX.WinActivate("EUKW - Prezentacja Księgi Wieczystej - Mozilla Firefox", "") == 1)
            {
                string numer_ksiegi = "000295527";
                string cyfra_kontrolna = "7";
                var notification = new System.Windows.Forms.NotifyIcon()
                {
                    Visible = true,
                    Icon = System.Drawing.SystemIcons.Information,
                    BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info,
                    BalloonTipTitle = "AutoIT - Ksiei Wieczyste",
                    BalloonTipText = "Trwa pobieranie danych z księgi wieczystej: SR2W/"+numer_ksiegi+'/'+cyfra_kontrolna,
                };
                notification.ShowBalloonTip(5000);

                //Przesuń i wpisz sąd
                AutoItX.MouseMove(2433, 464, 1);
                AutoItX.Sleep(500);
                AutoItX.MouseClick();
                AutoItX.Send("SR2W");
                AutoItX.Send("{TAB}");
                AutoItX.Sleep(500);

                //Wpisz numer księgi            
                AutoItX.Send(numer_ksiegi);
                AutoItX.Send("{TAB}");
                AutoItX.Sleep(500);

                //Wpisz cyfrę kontrolną
                AutoItX.Send(cyfra_kontrolna);
                AutoItX.Sleep(500);

                //przesuń nad google captcha
                AutoItX.MouseMove(2433, 546);
                AutoItX.Sleep(500);
                AutoItX.MouseClick();
                AutoItX.Sleep(4000);

                //wyszukaj księge
                AutoItX.MouseMove(2933, 685);
                AutoItX.Sleep(500);
                AutoItX.MouseClick();
                AutoItX.Sleep(6000);

                //Przeglądanie aktualnej treści KW
                AutoItX.MouseMove(2229, 797);
                AutoItX.Sleep(500);
                AutoItX.MouseClick();
                AutoItX.Sleep(5000);

                //Dzial III
                AutoItX.MouseMove(2512, 219);
                AutoItX.MouseClick();
                AutoItX.Sleep(600);
                AutoItX.Send("^u");
                AutoItX.Sleep(600);
                AutoItX.Send("^a");
                AutoItX.Sleep(500);
                AutoItX.Send("^c");
                AutoItX.Send("^{F4}");
                AutoItX.Sleep(1000);
                dzial_III = Clipboard.GetText();
                AutoItX.Sleep(500);

                //DZIAL II
                AutoItX.MouseMove(2349, 219);
                AutoItX.Sleep(500);
                AutoItX.MouseClick();
                AutoItX.Sleep(500);
                AutoItX.Send("^u");
                AutoItX.Sleep(600);
                AutoItX.Send("^a");
                AutoItX.Sleep(500);
                AutoItX.Send("^c");
                AutoItX.Send("^{F4}");
                AutoItX.Sleep(1000);
                dzial_II = Clipboard.GetText();
                AutoItX.Sleep(500);

                //DZIAL I-sp
                AutoItX.MouseMove(2176, 219);
                AutoItX.Sleep(500);
                AutoItX.MouseClick();
                AutoItX.Sleep(500);
                AutoItX.Send("^u");
                AutoItX.Sleep(600);
                AutoItX.Send("^a");
                AutoItX.Sleep(500);
                AutoItX.Send("^c");
                AutoItX.Send("^{F4}");
                AutoItX.Sleep(1000);
                dzial_Isp = Clipboard.GetText();
                AutoItX.Sleep(500);

                //DZIAL I-O
                AutoItX.MouseMove(1970, 219);
                AutoItX.Sleep(500);
                AutoItX.MouseClick();
                AutoItX.Sleep(500);
                AutoItX.Send("^u");
                AutoItX.Sleep(600);
                AutoItX.Send("^a");
                AutoItX.Sleep(500);
                AutoItX.Send("^c");
                AutoItX.Send("^{F4}");
                AutoItX.Sleep(1000);
                dzial_Io = Clipboard.GetText();
                AutoItX.Sleep(500);


                //Powrót do początku

                AutoItX.Send("{END}");
                AutoItX.MouseMove(1970, 946);
                AutoItX.Sleep(500);
                AutoItX.MouseClick();
                AutoItX.Sleep(1500);
                AutoItX.Send("{END}");
                AutoItX.Sleep(500);
                AutoItX.MouseMove(2662, 748);
                AutoItX.Sleep(500);
                AutoItX.MouseClick();

                //Powró do kryteriów

                AutoItX.Sleep(1500);

                AutoItX.Send("{END}");
                AutoItX.MouseMove(2661,747);

                AutoItX.Sleep(500);
                AutoItX.MouseClick();

                AutoItX.Sleep(1500);

                richTextBox1.AppendText(dzial_III);
                richTextBox2.AppendText(dzial_II);
                richTextBox3.AppendText(dzial_Isp);
                richTextBox4.AppendText(dzial_Io);

                StringBuilder numerKsiegi = new StringBuilder();
                numerKsiegi.Append("SR2W/");
                numerKsiegi.Append(numer_ksiegi);
                numerKsiegi.Append("/");
                numerKsiegi.Append(cyfra_kontrolna);

                dodajKsiegeDoBazy(numerKsiegi.ToString(), dzial_III, dzial_II, dzial_Isp, dzial_Io);

            }
            else
            {
                DialogResult dialogResult = MessageBox.Show("Czy uruchomić stronę EKW?", "Komunikat", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    var iPID = AutoItX.Run(@"C:\Program Files (x86)\Mozilla Firefox\firefox.exe https://przegladarka-ekw.ms.gov.pl/eukw_prz/KsiegiWieczyste/wyszukiwanieKW?komunikaty=true&kontakt=true&okienkoSerwisowe=false", "", 1);
                }
            }

            

        }

        private void button3_Click(object sender, EventArgs e)
        {
            var doc = new HtmlAgilityPack.HtmlDocument();
            doc.Load(@"c:\Users\Bartek\Desktop\wsad.html");
            HtmlNodeCollection name = doc.DocumentNode
    .SelectNodes("//table[@class='tbOdpis']/td");
            //td[@class='csDane']
            try
            {
                foreach (HtmlNode var in name)
                {
                    //richTextBox1.AppendText(System.Text.RegularExpressions.Regex.Replace(var.InnerHtml, @"<[^>]*>", String.Empty));
                    //richTextBox1.AppendText(var.InnerHtml);
                    //richTextBox1.AppendText("------NOWY WIERSZ------\n");
                    var test = System.Text.RegularExpressions.Regex.IsMatch(var.InnerHtml, @"^Lp.");
                    //richTextBox1.AppendText(test.ToString()+'\n');
                    if (test == true)
                    {
                        richTextBox1.AppendText(var.InnerHtml.ToString());
                    }
                }
            }
            catch(Exception x){
                MessageBox.Show(x.ToString());
            }
            
        }
        protected void dzial3_Click(object sender, EventArgs e)
        {
            richTextBox1.AppendText(dzial_III.ToString());
        }

        private void dodajKsiegeDoBazy(string nr_kw, string dz3,string dz2,string dz1sp,string dz1o)
        {
            try
            {
                SqlConnectionStringBuilder csb = new SqlConnectionStringBuilder();
                csb.ConnectionString = @"Data Source=(localdb)\MSSQLLocalDB;AttachDbFilename=C:\Users\Bartek\source\repos\Ksiegi wieczyste\Ksiegi wieczyste\KW.mdf;Integrated Security=True";
                SqlConnection conn = new SqlConnection(csb.ConnectionString);
                conn.Open();
                SqlCommand comm = new SqlCommand();
                string select = "SELECT * FROM kw";
                string insert = "INSERT INTO kw (kw,dz_3,dz_2,dz_1sp,dz_1o) VALUES (@nr_kw,@dz_3,@dz_2,@dz_1sp,@dz_1o)";
                comm.Connection = conn;
                comm.CommandText = insert;

                comm.Parameters.AddWithValue("@nr_kw", nr_kw);
                comm.Parameters.AddWithValue("@dz_3", dz3);
                comm.Parameters.AddWithValue("@dz_2", dz2);
                comm.Parameters.AddWithValue("@dz_1sp", dz1sp);
                comm.Parameters.AddWithValue("@dz_1o", dz1o);

                var result = comm.ExecuteNonQuery();

                comm.CommandText = select;
                SqlDataReader reader = comm.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        StringBuilder sb = new StringBuilder();
                        sb.Append(reader.GetInt32(0));
                        sb.Append(reader.GetValue(1));
                        sb.Append(reader.GetValue(3));
                        sb.Append("\n");
                        richTextBox1.AppendText(sb.ToString());
                    }
                }



                conn.Close();
            }
            catch (Exception x)
            {

                MessageBox.Show(x.ToString());
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {

            // wyszukajKsiege("SR2W/00002564/9");
            try
            {
                
                SqlConnectionStringBuilder csb = new SqlConnectionStringBuilder();
                csb.ConnectionString = @"Data Source=(localdb)\MSSQLLocalDB;AttachDbFilename="+ AppDomain.CurrentDomain.BaseDirectory.ToString()+@"data\kw.mdf"+"; Integrated Security=True";
                SqlConnection conn = new SqlConnection(csb.ConnectionString);
                conn.Open();
                string select = "SELECT * FROM kw";
                SqlCommand com = new SqlCommand();
                com.Connection = conn;
                com.CommandText = select;
                SqlDataReader reader = com.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        richTextBox1.AppendText(reader.GetString(1));
                        richTextBox1.AppendText("\n");
                    }
                }

                conn.Close();
            }
            catch (Exception x)
            {
                MessageBox.Show(x.ToString());
                throw;
            }
             var notification = new System.Windows.Forms.NotifyIcon()
            {
                Visible = true,
                Icon = System.Drawing.SystemIcons.Information,
                BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info,
                BalloonTipTitle = "AutoIT - Ksiei Wieczyste",
                BalloonTipText = "Trwa pobieranie danych z księgi wieczystej: ",
            };
            notification.ShowBalloonTip(5000);
            notification.Dispose();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Timer t1 = new Timer();
            t1.Interval = 50;
            t1.Tick += new EventHandler(timer1_Tick);
            t1.Enabled = true;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //richTextBox1.Clear();
           // richTextBox1.AppendText("X:"+Cursor.Position.X.ToString() + "Y:"+Cursor.Position.Y.ToString());
        }

        private void Form1_MouseClick(object sender, MouseEventArgs e)
        {
            
        }

        private void Form1_Deactivate(object sender, EventArgs e)
        {
            richTextBox1.Clear();
            richTextBox1.AppendText("X:" + Cursor.Position.X.ToString() + "Y:" + Cursor.Position.Y.ToString());
        }
    }
}
