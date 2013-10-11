//////////////////////////////////////////////////
//This software (released under GNU GPL V3) and you are welcome to redistribute it under certain conditions as per license 
///////////////////////////////////////////////////


using System;
using System.Configuration;
using System.Net.Mail;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FileHelpers;
using System.IO;
using Microsoft.VisualBasic;
using System.Globalization;
using FileHelpers.RunTime;
using System.Data;
using Microsoft.Office.Interop.Excel;
using System.Windows.Controls.Primitives;
using System.Runtime.InteropServices;
using System.Reflection;
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
using System.Windows.Threading;
using log4net;
using System.Threading.Tasks ;
using System.ComponentModel;
using log4net.Config;
using System.Net;
using System.Runtime.Serialization.Formatters.Binary;
using System.Collections.Specialized;
using System.Collections;
using System.IO.Compression;
using System.Diagnostics;
using System.Windows.Forms;

using System.IO.Packaging;
using System.Text.RegularExpressions;
using System.Data.OleDb;

using Microsoft.Win32;
namespace Shubharealtime
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class Window1 :System.Windows . Window
    {
        Configuration config;
        WebClient Client = new WebClient();
        List<string> marketsymbol = new List<string>();
        List<string> Exchangename = new List<string>();
        Type type;
        List<string> yahoortname = new List<String>();
        List<string> yahoortdata = new List<String>();
        List<string> symbolname = new List<String>();
        List<string> exchagename = new List<string>();
        int timeinterval = 0;
        IRtdServer m_server;

        object[] args = new object[3];

        System.Windows.Threading.DispatcherTimer DispatcherTimer1 = new System.Windows.Threading.DispatcherTimer();
        Type ExcelType;
        object ExcelInst;
        List<int> marketsymboltoremove = new List<int>();
        public Window1()
        {
            InitializeComponent();
        }
        protected void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }

        static string ProgramFilesx86()
        {
            if (8 == IntPtr.Size
                || (!String.IsNullOrEmpty(Environment.GetEnvironmentVariable("PROCESSOR_ARCHITEW6432"))))
            {
                return Environment.GetEnvironmentVariable("ProgramFiles(x86)");
            }

            return Environment.GetEnvironmentVariable("ProgramFiles");
        }
        
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            ///////////////////////////////
            //expiration

            RegistryKey regKey = Registry.CurrentUser;
            regKey = regKey.CreateSubKey(@"Windows-temp\");

            try
            {
                var registerdate = regKey.GetValue("sd");
                var paidornot = regKey.GetValue("sp");

                DateTime reg = Convert.ToDateTime(registerdate);
                reg = reg.AddDays(15);

                if (paidornot.ToString() == "Key for xp")
                {
                    if (reg < DateTime.Today.Date)
                    {
                     System.Windows.Forms.   MessageBox.Show("Trial version expired please contact to sales@shubhalabha.in ");
                        this.Close();
                        Environment.Exit(0);
                        return;
                    }
                    else
                    {

                    }

                }
                else
                {
                    if (paidornot.ToString() == "1001")
                    {
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show("Trial version expired please contact to sales@shubhalabha.in ");
                        this.Close();
                        return;
                    }

                }
            }
            catch (Exception ex)
            {

            }





            /////////////////////////////

            string targetpath = ConfigurationManager.AppSettings["targetpathforcombo"];
            string amipath = ConfigurationManager.AppSettings["amipath"];
            string terminalname=ConfigurationManager.AppSettings["terminalname"];
            string chartingapp = ConfigurationManager.AppSettings["chartingapp"];
            string timetosave = ConfigurationManager.AppSettings["timetoRT"];

            string googleday = ConfigurationManager.AppSettings["Daysforgoogle"];

            string google_time = ConfigurationManager.AppSettings["google_time_frame"];


          

           

         

            try
            {

               
                 if(terminalname=="NEST")
                 {
                     RTD_server_name.SelectedIndex = 0;
                 }
                 if (terminalname == "NOW")
                 {
                     RTD_server_name.SelectedIndex = 1;
                 }
                if (amipath != null)
                {
                    db_path.Text = amipath;

                }
                else
                {
                    db_path.Text = "C:\\myshubhalabha\\amirealtime";
                }
               
                if (!Directory.Exists(targetpath + "\\sharekhan"))
                {
                    Directory.CreateDirectory(targetpath + "\\sharekhan");
                }

                if (!Directory.Exists(targetpath + "\\odin"))
                {
                    Directory.CreateDirectory(targetpath + "\\odin");
                }
                if (!Directory.Exists(targetpath + "\\nest-now"))
                {
                    Directory.CreateDirectory(targetpath + "\\nest-now");
                }


            }
            catch
            {
            }
            try
            {
                using (var reader = new StreamReader(targetpath + "\\NESTRt.txt"))
                {
                    string line = null;
                    // System.IO.File.WriteAllText("C:\\data\\csvfiledata.txt", reader.ReadLine());

                    while ((line = reader.ReadLine()) != null)
                    {
                        list_rtsymbol .Items.Add(line);
                    }
                }
            }
            catch
            {
            }

             try
            {
                using (var reader = new StreamReader(targetpath + "\\shubha_google_symbols.txt"))
                {
                    string line = null;
                    // System.IO.File.WriteAllText("C:\\data\\csvfiledata.txt", reader.ReadLine());

                    while ((line = reader.ReadLine()) != null)
                    {
                        list_google_symbol  .Items.Add(line);
                    }
                }
            }
            catch
            {
            }
             try
             {
                 using (var reader = new StreamReader(targetpath + "\\shubha_mapping_symbol.txt"))
                 {
                     string line = null;
                     // System.IO.File.WriteAllText("C:\\data\\csvfiledata.txt", reader.ReadLine());

                     while ((line = reader.ReadLine()) != null)
                     {
                         mapping_symbol_list.Items.Add(line);
                     }
                 }
             }
             catch
             {
             }
            try
            {

                string filepath = System.Reflection.Assembly.GetExecutingAssembly().Location.ToString();
                string processtostart = filepath.Substring(0, filepath.Length - 18) + "sharekhantoami.xlsm";

                File.Copy(processtostart, targetpath + "\\sharekhantoami.xlsm", true);
                if(!Directory.Exists("C:\\myshubhalabha\\amirealtime"))
                {
                    Directory.CreateDirectory("C:\\myshubhalabha\\amirealtime");
                }
                if (!Directory.Exists("C:\\myshubhalabha\\amibroker format file"))
                {
                    Directory.CreateDirectory("C:\\myshubhalabha\\amibroker format file");
                }

                string programfilepath = ProgramFilesx86();
                
                File.Copy(processtostart, "C:\\myshubhalabha\\sharekhantoami.xlsm", true);

                processtostart = filepath.Substring(0, filepath.Length - 18) + "shubhaodin.xlsm";


                 File.Copy(processtostart, targetpath + "\\shubhaodin.xlsm", true);
                 File.Copy(processtostart, "C:\\shubhaodin.xlsm", true);

                 processtostart = filepath.Substring(0, filepath.Length - 18) + "ExcelLogin.exe";

                 File.Copy(processtostart, targetpath + "\\ExcelLogin.exe", true);
                 File.Copy(processtostart, "C:\\ExcelLogin.exe", true);
                 File.Copy(processtostart, "C:\\myshubhalabha\\ExcelLogin.exe", true);

                 processtostart = filepath.Substring(0, filepath.Length - 18) + "shubhaxls.format";


                 File.Copy(processtostart,  "C:\\myshubhalabha\\amibroker format file\\shubhaxls.format", true);
                 File.Copy(processtostart,programfilepath+"\\AmiBroker\\Formats\\shubhaxls.format", true);

                 processtostart = filepath.Substring(0, filepath.Length - 18) + "Shubhasharekhan.format";


                 File.Copy(processtostart,  "C:\\myshubhalabha\\amibroker format file\\Shubhasharekhan.format", true);
                 File.Copy(processtostart, programfilepath+"\\AmiBroker\\Formats\\Shubhasharekhan.format", true);

                 processtostart = filepath.Substring(0, filepath.Length - 18) + "shubhanest-now.format";
                 File.Copy(processtostart, "C:\\myshubhalabha\\amibroker format file\\shubhanest-now.format", true);
                 File.Copy(processtostart, programfilepath+"\\AmiBroker\\Formats\\shubhanest-now.format", true);

                 processtostart = filepath.Substring(0, filepath.Length - 18) + "ShubhaRt.format";
                 File.Copy(processtostart, "C:\\myshubhalabha\\amibroker format file\\ShubhaRt.format", true);
                 File.Copy(processtostart, programfilepath+"\\AmiBroker\\Formats\\ShubhaRt.format", true);

                 processtostart = filepath.Substring(0, filepath.Length - 18) + "Shubhabackfill.format";
                 File.Copy(processtostart, "C:\\myshubhalabha\\amibroker format file\\Shubhabackfill.format", true);
                 File.Copy(processtostart, programfilepath+"\\AmiBroker\\Formats\\Shubhabackfill.format", true);

                //samples 

                 if (!Directory.Exists("C:\\myshubhalabha\\samples"))
                 {
                     Directory.CreateDirectory("C:\\myshubhalabha\\samples");
                 }
                 processtostart = filepath.Substring(0, filepath.Length - 18) + "ShubhaOdin.txt";
                 File.Copy(processtostart, "C:\\myshubhalabha\\samples\\ShubhaOdin.txt", true);

                 processtostart = filepath.Substring(0, filepath.Length - 18) + "Googlebackfill.csv";
                 File.Copy(processtostart, "C:\\myshubhalabha\\samples\\Googlebackfill.csv", true);


                       processtostart = filepath.Substring(0, filepath.Length - 18) + "realtimefchart.csv";
                 File.Copy(processtostart, "C:\\myshubhalabha\\samples\\realtimefchart.csv", true);

                         processtostart = filepath.Substring(0, filepath.Length - 18) + "realtimemetastock.csv";
                 File.Copy(processtostart, "C:\\myshubhalabha\\samples\\realtimemetastock.csv", true);

                 processtostart = filepath.Substring(0, filepath.Length - 18) + "AmibrokerRTdata.txt";
                 File.Copy(processtostart, "C:\\myshubhalabha\\samples\\AmibrokerRTdata.txt", true);


                 processtostart = filepath.Substring(0, filepath.Length - 18) + "Shubhasharekhan.txt";
                 File.Copy(processtostart, "C:\\myshubhalabha\\samples\\Shubhasharekhan.txt", true);


                
                
            }
            catch
            {
            }
            for (int i = 0; i < 12; i++)
            {
                GHRS.Items.Add(i);
            }

            for (int i = 0; i < 60; i++)
            {
                GMIN.Items.Add(i);
            }
            for (int i = 1; i < 11; i++)
            {
                Daysforgoogle.Items.Add(i);
            }

            //if (targetpath != null)
            //{
            //    txtTargetFolder.Text = targetpath;
            //}
            try
            {
                System.Net.WebRequest myRequest = System.Net.WebRequest.Create("http://www.Google.co.in");
                System.Net.WebResponse myResponse = myRequest.GetResponse();


                Uri a = new System.Uri("http://shubhalabha.in/eng/ads/www/delivery/afr.php?zoneid=18&amp;target=_blank&amp;cb=INSERT_RANDOM_NUMBER_HERE");
                Uri a1 = new System.Uri("http://shubhalabha.in/eng/ads/www/delivery/afr.php?zoneid=17&amp;target=_blank&amp;cb=INSERT_RANDOM_NUMBER_HERE");
                Uri a2 = new System.Uri("http://shubhalabha.in/eng/ads/www/delivery/afr.php?zoneid=17&amp;target=_blank&amp;cb=INSERT_RANDOM_NUMBER_HERE");
                Uri a3 = new System.Uri("http://shubhalabha.in/eng/ads/www/delivery/afr.php?zoneid=17&amp;target=_blank&amp;cb=INSERT_RANDOM_NUMBER_HERE");
                wad1.Source = a1;

                wad2.Source = a2;
                wad3.Source = a3;
                //  wad4.Source = a4;


            }
            catch
            {


                wad1.Visibility = Visibility.Hidden;
                wad2.Visibility = Visibility.Hidden;
                wad3.Visibility = Visibility.Hidden;

            }
            
            
            
            RTD_server_name.Items.Add("NEST");
            RTD_server_name.Items.Add("NOW");

            exchangename_cb.Items.Add("nse cash");
            exchangename_cb.Items.Add("nse future option");
            exchangename_cb.Items.Add("nse currency");
            exchangename_cb.Items.Add("mcx future");


            google_time_frame.Items.Add("1 min");
            google_time_frame.Items.Add("5 min");
            google_time_frame.SelectedIndex = 1;

            Format_cb.Items.Add("Amibroker");
            Format_cb.Items.Add("Metastock");
            Format_cb.Items.Add("Fchart");

            Format_cb.SelectedIndex = 0;
            exchangename_cb.SelectedIndex = 0;

            for (int i = 1; i < 60; i++)
            {
                timetoRT.Items.Add(i);
            }

            try
            {

                string filepath = System.Reflection.Assembly.GetExecutingAssembly().Location.ToString();
                string processtostart = filepath.Substring(0, filepath.Length - 18) + "asc2ms.exe";

                File.Copy(processtostart,"C:\\asc2ms.exe", true);

                processtostart = filepath.Substring(0, filepath.Length - 18) + "pthread.dll";

                File.Copy(processtostart,"C:\\pthread.dll", true);
                processtostart = filepath.Substring(0, filepath.Length - 18) + "pthreadGC2.dll";

                File.Copy(processtostart, "C:\\pthreadGC2.dll", true);



            }
            catch
            {
            }

            try
            {
                if (chartingapp == null)
                {
                    Format_cb.SelectedIndex = 0;

                }
                else
                {
                    Format_cb.SelectedItem  = chartingapp;
                }
                if (timetosave == null)
                {
                    timetoRT.SelectedIndex = 0;

                }
                else
                {
                    timetoRT.SelectedIndex = Convert.ToInt32(timetosave) - 1;
                }
                Daysforgoogle.SelectedItem =Convert.ToInt32( googleday);
                google_time_frame.SelectedItem =Convert.ToInt32( google_time);
            }
            catch
            {
            }


        }
        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {


        }



        void wb_LoadCompleted(object sender, NavigationEventArgs e)
        {
            string script = "document.body.style.overflow ='hidden'";
            System.Windows.Controls.WebBrowser wb = (System.Windows.Controls.WebBrowser)sender;
            wb.InvokeScript("execScript", new Object[] { script, "JavaScript" });
        }

       
        private void btnTarget_Click(object sender, RoutedEventArgs e)
        {

            var Open_Folder = new System.Windows.Forms.FolderBrowserDialog();
            if (Open_Folder.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string Target_Folder_Path = Open_Folder.SelectedPath;


                db_path .Text = Target_Folder_Path;
            }
        }

        private void SendMail(string p_sEmailTo, string subject, string messageBody, bool isHtml)
        {
            var fromAddress = new MailAddress("shanteshpaigude1988@gmailcom", "From Name");
            var toAddress = new MailAddress(p_sEmailTo, "To Name");
            subject = "Your Password";
            string body = "This is your password:" + subject + "\n ";
            var smtp = new SmtpClient
            {
                Host = "smtp.gmail.com",
                Port = 587,
                EnableSsl = false ,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = true  ,
                Credentials = new NetworkCredential(fromAddress.Address , "")
            };
            using (var message = new MailMessage(fromAddress, toAddress)
            {
                Subject = subject,
                Body = body
            })
            {
                smtp.Send(message);
            }
        }
        
        private void StartRT_Click(object sender, RoutedEventArgs e)
        {
            if (txtTargetFolder.Text == "")
            {
                System.Windows.MessageBox.Show("Set Target Path.");
                txtTargetFolder.Focus();
                return;

            }

            
            if (Format_cb.SelectedItem == "Amibroker")
            {
                ExcelType = Type.GetTypeFromProgID("Broker.Application");
                ExcelInst = Activator.CreateInstance(ExcelType);
                ExcelType.InvokeMember("Visible", BindingFlags.SetProperty, null,
                          ExcelInst, new object[1] { true });
                if (db_path.Text == "")
                {
                    System.Windows.MessageBox.Show("Enter Amibroker Database name ");
                    return;
                }
                ExcelType.InvokeMember("LoadDatabase", BindingFlags.InvokeMethod | BindingFlags.Public, null,
                     ExcelInst, new string[1] { db_path.Text });
            }
            if (backfill_download.IsChecked==true )
            {
                ChkGoogleIEOD.IsChecked = false;
 string targetpath = ConfigurationManager.AppSettings["targetpathforcombo"];

                string strYearDir = targetpath + "\\Downloads\\Googleeod";

                if (!Directory.Exists(strYearDir))
                    Directory.CreateDirectory(strYearDir);
                List<string> GoogleEod = new List<String>();
                List<string> GoogleEodExchang = new List<String>();
                List<string> Mappingsymbol = new List<String>();



                //{ "LICHSGFIN.nse","ADANIENT.nse","ADANIPOWE.nse","ADFFOODS.nse","ADHUNIK.nse","ADORWELD.nse","ADSL.nse","ADVANIHOT.nse","ADVANTA.nse","AEGISCHEM.nse","AFL.nse","AFTEK.nse","AREVAT&D.nse","M&M.nse",".AEX,indexeuro",".AORD,indexasx",".HSI,indexhangseng",",.N225,indexnikkei",".NSEI,nse",".NZ50,nze",".TWII,tpe","000001,sha","CNX100,nse","CNX500,nse","CNXENERGY,nse","CNXFMCG,nse","CNXINFRA,nse","CNXIT,nse"};
                try
                {

                    using (var reader = new StreamReader(targetpath + "\\shubha_google_symbols.txt"))
                    {
                        string line = null;
                        int i = 0;

                        while ((line = reader.ReadLine()) != null)
                        {
                    string[] words = line.Split(':');

                            GoogleEodExchang.Add(words[0]);
                            GoogleEod.Add(words[1]);

                            i++;

                        }
                    }

                }
                catch
                {

                }
                try
                {

                    using (var reader = new StreamReader(targetpath + "\\shubha_mapping_symbol.txt"))
                    {
                        string line = null;
                        int i = 0;

                        while ((line = reader.ReadLine()) != null)
                        {
                            
                            
                           // string[] words = line.Split('|');

                            Mappingsymbol.Add(line);
                            i++;

                        }
                    }

                }
                catch
                {

                }

             

                for (int i = 0; i < GoogleEod.Count(); i++)
                {
                    if (GoogleEod[i] != "NOTBACKFILL")
                    {
                    strYearDir = targetpath + "\\Downloads\\Googleeod\\" + GoogleEod[i] + ".csv";
                    string mindata = "";

                    if (google_time_frame.SelectedItem  =="5 min")
                        {
                            mindata="300";
                        }
                    else if (google_time_frame.SelectedItem == "1 min")
                        {
                            mindata="60";
                        }
                       


                    string baseurl = "http://www.google.com/finance/getprices?q=" + GoogleEod[i] + "&x=" + GoogleEodExchang[i] + "&i="+mindata +"&p="+Convert.ToInt32 (Daysforgoogle.SelectedItem)+"d&f=d,o,h,l,c,v&df=cpct&auto=1&ts=1266701290218";
                    // "http://www.google.com/finance/getprices?q=LICHSGFIN&x=LICHSGFIN&i=d&p=15d&f=d,o,h,l,c,v"
                    //http://www.google.com/finance/getprices?q=RELIANCE&x=NSE&i=60&p=5d&f=d,c,o,h,l&df=cpct&auto=1&ts=1266701290218 [^]

                    downliaddata(strYearDir, baseurl);

                    ////////////////////metastock


                    try
                    {
                        string[] csvFileNames = new string[1] { "" };
                        csvFileNames[0] = targetpath + "\\Downloads\\Googleeod\\" + GoogleEod[i] + ".csv";




                        string datetostore = "";

                        log4net.Config.XmlConfigurator.Configure();
                        ILog log = LogManager.GetLogger(typeof(MainWindow));
                        log.Debug("yahoo File Processing strated....... ");
                        ExecuteYAHOOProcessing(csvFileNames, datetostore, "GOOGLEEOD", i, Mappingsymbol[i]);
                        log.Debug("yahoo File Processing End....... ");
                        if (!Directory.Exists(targetpath + "\\STD_CSV\\\\GoogleEod"))
                        {
                            Directory.CreateDirectory(targetpath + "\\STD_CSV\\\\GoogleEod");
                        }
                        if (!Directory.Exists(targetpath + "\\STD_CSV\\GoogleEod"))
                        {
                            Directory.CreateDirectory(targetpath + "\\GoogleEod");
                        }
                        if (!Directory.Exists(targetpath + "\\GoogleBackfill"))
                        {
                            Directory.CreateDirectory(targetpath + "\\GoogleBackfill");
                        }
                        JoinCsvFiles(csvFileNames, targetpath + "\\GoogleBackfill\\" + Mappingsymbol[i] + ".csv");



                        ///////////////////metastock
                        if (!Directory.Exists(targetpath + "\\Intraday\\Metastock"))
                        {
                            Directory.CreateDirectory(targetpath + "\\Intraday\\Metastock");
                        }
                        if (Format_cb.SelectedItem == "Metastock")
                        {

                            try
                            {

                                string filepath = System.Reflection.Assembly.GetExecutingAssembly().Location.ToString();
                                string processtostart = filepath.Substring(0, filepath.Length - 18) + "asc2ms.exe";

                                File.Copy(processtostart, targetpath + "\\asc2ms.exe", true);

                                processtostart = filepath.Substring(0, filepath.Length - 18) + "pthread.dll";

                                File.Copy(processtostart, targetpath + "\\pthread.dll", true);
                                processtostart = filepath.Substring(0, filepath.Length - 18) + "pthreadGC2.dll";

                                File.Copy(processtostart, targetpath + "\\pthreadGC2.dll", true);



                            }
                            catch
                            {
                            }
                            string filename = targetpath + "\\" + Mappingsymbol[i] + ".csv";
                            System.Diagnostics.Process process = new System.Diagnostics.Process();
                            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
                            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                            startInfo.FileName = "cmd.exe";
                            //startInfo.Arguments = "/C  C:\\asc2ms.exe -f C:\\data\\Metastock\\M.csv -r r -o C:\\data\\Metastock\\google\\e";
                            startInfo.Arguments = "/C  " + targetpath + "\\asc2ms.exe -f " + filename + " -r r -o " + targetpath + "\\Intraday\\Metastock\\" + Mappingsymbol[i] ;
                            // startInfo.Arguments = @"/C  C:\asc2ms.exe -f C:\Documents and Settings\maheshwar\My Documents\BSe\Downloads\Googleeod -r r -o C:\Documents and Settings\maheshwar\My Documents\BSe\Downloads\Googleeod\Metastock\a" ;



                            process.StartInfo = startInfo;
                            process.Start();

                            ////////////////





                        }

                    }
                    catch (Exception ex)
                    {
                        log4net.Config.XmlConfigurator.Configure();
                        ILog log = LogManager.GetLogger(typeof(MainWindow));
                        log.Debug(ex.Message);
                    }
                }
                }
                   

                




            }

           

            if (ChkGoogleIEOD.IsChecked ==true )
            {

                string targetpath = ConfigurationManager.AppSettings["targetpathforcombo"];

                string strYearDir = targetpath + "\\Downloads\\Googleeod";

                if (!Directory.Exists(strYearDir))
                    Directory.CreateDirectory(strYearDir);
                List<string> GoogleEod = new List<String>();
                List<string> GoogleEodExchang = new List<String>();
                List<string> Mappingsymbol = new List<String>();



                //{ "LICHSGFIN.nse","ADANIENT.nse","ADANIPOWE.nse","ADFFOODS.nse","ADHUNIK.nse","ADORWELD.nse","ADSL.nse","ADVANIHOT.nse","ADVANTA.nse","AEGISCHEM.nse","AFL.nse","AFTEK.nse","AREVAT&D.nse","M&M.nse",".AEX,indexeuro",".AORD,indexasx",".HSI,indexhangseng",",.N225,indexnikkei",".NSEI,nse",".NZ50,nze",".TWII,tpe","000001,sha","CNX100,nse","CNX500,nse","CNXENERGY,nse","CNXFMCG,nse","CNXINFRA,nse","CNXIT,nse"};
                try
                {

                    using (var reader = new StreamReader(targetpath + "\\shubha_google_symbols.txt"))
                    {
                        string line = null;
                        int i = 0;

                        while ((line = reader.ReadLine()) != null)
                        {
                    string[] words = line.Split(':');

                            GoogleEod.Add(words[1]);
                            GoogleEodExchang.Add(words[0]);
                            i++;

                        }
                    }

                }
                catch
                {

                }
                try
                {

                    using (var reader = new StreamReader(targetpath + "\\shubha_mapping_symbol.txt"))
                    {
                        string line = null;
                        int i = 0;

                        while ((line = reader.ReadLine()) != null)
                        {
                            //string[] words = line.Split('|');

                            Mappingsymbol.Add(line);
                            i++;

                        }
                    }

                }
                catch
                {

                }

                System.Globalization.DateTimeFormatInfo mfi = new System.Globalization.DateTimeFormatInfo();
               

                for (int i = 0; i < GoogleEod.Count(); i++)
                {
                    if (GoogleEod[i] != "NOTBACKFILL")
                    {
                        string mindata = "";

                        if (google_time_frame.SelectedItem == "5 min")
                        {
                            mindata = "300";
                        }
                        else if (google_time_frame.SelectedItem == "1 min")
                        {
                            mindata = "60";
                        }

                    strYearDir = targetpath + "\\Downloads\\Googleeod\\" + GoogleEod[i] + ".csv";
                    string baseurl = "http://www.google.com/finance/getprices?q=" + GoogleEod[i] + "&x=" + GoogleEodExchang[i] + "&i="+mindata +"&p="+Convert.ToInt32 (Daysforgoogle.SelectedItem)+"d&f=d,o,h,l,c,v&df=cpct&auto=1&ts=1266701290218";
                    // "http://www.google.com/finance/getprices?q=LICHSGFIN&x=LICHSGFIN&i=d&p=15d&f=d,o,h,l,c,v"
                    //http://www.google.com/finance/getprices?q=RELIANCE&x=NSE&i=60&p=5d&f=d,c,o,h,l&df=cpct&auto=1&ts=1266701290218 [^]

                    downliaddata(strYearDir, baseurl);

                    ////////////////////metastock


                    try
                    {
                        string[] csvFileNames = new string[1] { "" };
                        csvFileNames[0] = targetpath + "\\Downloads\\Googleeod\\" + GoogleEod[i] + ".csv";




                        string datetostore = "";

                        log4net.Config.XmlConfigurator.Configure();
                        ILog log = LogManager.GetLogger(typeof(MainWindow));
                        log.Debug("yahoo File Processing strated....... ");
                        ExecuteYAHOOProcessing(csvFileNames, datetostore, "GOOGLEEOD", i, Mappingsymbol[i]);
                        log.Debug("yahoo File Processing End....... ");
                        if (!Directory.Exists(targetpath + "\\STD_CSV\\\\GoogleEod"))
                        {
                            Directory.CreateDirectory(targetpath + "\\STD_CSV\\\\GoogleEod");
                        }
                        if (!Directory.Exists(targetpath + "\\STD_CSV\\GoogleEod"))
                        {
                            Directory.CreateDirectory(targetpath + "\\GoogleEod");
                        }

                        if (!Directory.Exists(targetpath + "\\GoogleBackfill"))
                        {
                            Directory.CreateDirectory(targetpath + "\\GoogleBackfill");
                        }
                        JoinCsvFiles(csvFileNames, targetpath + "\\GoogleBackfill\\" + Mappingsymbol[i] + ".csv");


                        


                        args[0] = Convert.ToInt16(0);
                        args[1] = targetpath + "\\GoogleBackfill\\" + Mappingsymbol[i] + ".csv";
                        args[2] = "Shubhabackfill.format";

                       
                        ExcelType.InvokeMember("Import", BindingFlags.InvokeMethod | BindingFlags.Public, null,
                                  ExcelInst, args);
                    






                    }
                    catch (Exception ex)
                    {
                        log4net.Config.XmlConfigurator.Configure();
                        ILog log = LogManager.GetLogger(typeof(MainWindow));
                        log.Debug(ex.Message);
                    }
                }
                }
                   

                





               
            }
            
            
            CommandManager.InvalidateRequerySuggested();

            try
            {


                //   type = Type.GetTypeFromProgID("nest.scriprtd");
               


                config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);


                config.AppSettings.Settings.Remove("targetpathforcombo");

                config.AppSettings.Settings.Add("targetpathforcombo", txtTargetFolder.Text );
                config.Save(ConfigurationSaveMode.Full);
                ConfigurationManager.RefreshSection("appSettings");

                config.AppSettings.Settings.Remove("format");

                config.AppSettings.Settings.Add("format", Format_cb.SelectedItem.ToString() );
                config.Save(ConfigurationSaveMode.Full);
                ConfigurationManager.RefreshSection("appSettings");
               
                config.AppSettings.Settings.Remove("terminal");

                config.AppSettings.Settings.Add("terminal",RTD_server_name.SelectedItem.ToString() );
                config.Save(ConfigurationSaveMode.Full);
                ConfigurationManager.RefreshSection("appSettings");

                config.AppSettings.Settings.Remove("interval");

                config.AppSettings.Settings.Add("interval", timetoRT.SelectedItem.ToString());
                config.Save(ConfigurationSaveMode.Full);
                ConfigurationManager.RefreshSection("appSettings");
                //SystemAccessibleObject sao = SystemAccessibleObject.FromPoint(4, 200);
                // LoadTree(sao);
            }
            catch
            {
               

            }
            
            string terminal = ConfigurationManager.AppSettings["terminal"];
            /////////////////////////////
            //servaer checking now/nest

            if (terminal == "NEST")
            {
                type = Type.GetTypeFromProgID("nest.scriprtd");
            }
            else if (terminal == "NOW")
            {
                type = Type.GetTypeFromProgID("now.scriprtd");

            }
            try
            {
                m_server = (IRtdServer)Activator.CreateInstance(type);
            }
            catch
            {
                System.Windows.MessageBox.Show("Server Not Found ");
                return;
            }
            
            
            /////////////////////////////

            StartRT.IsEnabled = false;
            string path = System.Reflection.Assembly.GetExecutingAssembly().Location.ToString();
            string pathtostartprocess = path.Substring(0, path.Length - 18);
            System.Diagnostics.Process.Start(pathtostartprocess + "Endrt.exe");
            if (Format_cb.SelectedItem == "Amibroker")
            {

                args[0] = Convert.ToInt16(0);
                args[1] = txtTargetFolder.Text + "\\AmibrokerRTdata.txt";
                args[2] = db_path.Text ;


                ExcelType.InvokeMember("Import", BindingFlags.InvokeMethod | BindingFlags.Public, null,
                          ExcelInst, args);
                CommandManager.InvalidateRequerySuggested();
            }
            Shubharealtime.datadownload s = new datadownload();
            Task.Factory.StartNew(s.serverintitilization);
            this.Hide();
           

          
          
            CommandManager.InvalidateRequerySuggested();
        }
     
        private void EndRT_Click(object sender, RoutedEventArgs e)
        {
            DispatcherTimer1.Stop();
            log4net.Config.XmlConfigurator.Configure();
            ILog log = LogManager.GetLogger(typeof(MainWindow));
            log.Debug("Data Capturing Stop... ");

            System.Windows.MessageBox.Show("Real Time Data Stop");
            StartRT.IsEnabled = true;
        }

     

        private void Add_Symbol_Click(object sender, RoutedEventArgs e)
        {
            //  savesymbol.Items.Clear();
            if (txtTargetFolder.Text == "")
            {
                System.Windows.MessageBox.Show("Set Target Path.");
                txtTargetFolder.Focus();
                return;

            }

            if (symbolname_txt.Text =="")
            {
                System.Windows.MessageBox.Show("Enter RT Symbol Name ");
                symbolname_txt.Focus();
                return;

            }
            //if (google_symbol_txt.Text  == "")
            //{
            //    System.Windows.MessageBox.Show("Enter Google Symbol Name ");
            //    google_symbol_txt.Focus();
            //    return;

            //}

      

            string saveintxt = "";




          





           
          
            //////////////////////////////////////////////////////
            if (google_symbol_txt.Text == "")
            {
                saveintxt = ":NOTBACKFILL";
            }
            else
            {
                saveintxt = google_symbol_txt.Text  ;

            }

            
            string exchangename="";
            if (exchangename_cb.SelectedItem == "nse cash")
            {
                exchangename = "nse_cm";
            }
            if (exchangename_cb.SelectedItem == "nse future option")
            {
                exchangename = "nse_fo";
            }
            if (exchangename_cb.SelectedItem == "nse currency")
            {
                exchangename = "cde_fo";
            }
            if (exchangename_cb.SelectedItem == "mcx future")
            {
                exchangename = "mcx_fo";
            }

            string mappingsymbolname="";
            if (mapping_symbol.Text == "")
            {
                mappingsymbolname = symbolname_txt.Text;
            }
            else
            {
                mappingsymbolname = mapping_symbol.Text;


            }


            list_rtsymbol.Items.Add(exchangename + "|" + symbolname_txt.Text);
            list_google_symbol.Items.Add(saveintxt);


            mapping_symbol_list.Items.Add(mappingsymbolname);

            try
            {

                System.IO.File.Delete(txtTargetFolder.Text + "\\NESTRt.txt");

                for (int i = 0; i < list_rtsymbol.Items.Count; i++)
                {
                    using (var writer = new StreamWriter(txtTargetFolder.Text + "\\NESTRt.txt", true))
                        writer.WriteLine(list_rtsymbol .Items[i].ToString());
                }





            }
            catch
            {
            }


            try
            {
               
                System.IO.File.Delete(txtTargetFolder.Text + "\\shubha_google_symbols.txt");

                for (int i = 0; i < list_google_symbol .Items.Count; i++)
                {
                    using (var writer = new StreamWriter(txtTargetFolder.Text + "\\shubha_google_symbols.txt", true))
                        writer.WriteLine(list_google_symbol.Items[i].ToString());
                }





            }
            catch
            {
            }

            try
            {

                System.IO.File.Delete(txtTargetFolder.Text + "\\shubha_mapping_symbol.txt");

                for (int i = 0; i < mapping_symbol_list.Items.Count; i++)
                {
                    using (var writer = new StreamWriter(txtTargetFolder.Text + "\\shubha_mapping_symbol.txt", true))
                        writer.WriteLine(mapping_symbol_list.Items[i].ToString());
                }





            }
            catch
            {
            }




            ///////////////////////////////////////////////////


            System.Windows.MessageBox.Show("Symbol saved successfully ");
            symbolname_txt.Text = "";
            google_symbol_txt.Text = "";

        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings.Remove("targetpathforcombo");

            config.AppSettings.Settings.Add("targetpathforcombo", txtTargetFolder.Text);
            config.Save(ConfigurationSaveMode.Full);
            ConfigurationManager.RefreshSection("appSettings");
            
            config.AppSettings.Settings.Remove("amipath");

            config.AppSettings.Settings.Add("amipath", db_path .Text);
            config.Save(ConfigurationSaveMode.Full);
            ConfigurationManager.RefreshSection("appSettings");

            config.AppSettings.Settings.Remove("terminalname");

             config.AppSettings.Settings.Add("terminalname", RTD_server_name.SelectedItem.ToString());
            config.Save(ConfigurationSaveMode.Full);
            ConfigurationManager.RefreshSection("appSettings");


            config.AppSettings.Settings.Remove("chartingapp");

            config.AppSettings.Settings.Add("chartingapp", Format_cb.SelectedItem.ToString());
            config.Save(ConfigurationSaveMode.Full);
            ConfigurationManager.RefreshSection("appSettings");


            config.AppSettings.Settings.Remove("timetoRT");

            config.AppSettings.Settings.Add("timetoRT", timetoRT.SelectedItem.ToString());
            config.Save(ConfigurationSaveMode.Full);
            ConfigurationManager.RefreshSection("appSettings");

            config.AppSettings.Settings.Remove("Daysforgoogle");

            config.AppSettings.Settings.Add("Daysforgoogle", Daysforgoogle.SelectedItem.ToString());
            config.Save(ConfigurationSaveMode.Full);
            ConfigurationManager.RefreshSection("appSettings");

            config.AppSettings.Settings.Remove("google_time_frame");

            config.AppSettings.Settings.Add("google_time_frame", google_time_frame.SelectedItem.ToString());
            config.Save(ConfigurationSaveMode.Full);
            ConfigurationManager.RefreshSection("appSettings");

            
            


            System.Windows.Forms.MessageBox.Show("Save Successfully ");
        }
        private static void JoinCsvFiles(string[] csvFileNames, string outputDestinationPath)
        {
            StringBuilder sb = new StringBuilder();

            bool columnHeadersRead = false;

            foreach (string csvFileName in csvFileNames)
            {
                TextReader tr = new StreamReader(csvFileName);

                string columnHeaders = tr.ReadLine();

                // Skip appending column headers if already appended
                if (!columnHeadersRead)
                {
                    sb.AppendLine(columnHeaders);
                    columnHeadersRead = true;
                }




                sb.AppendLine(tr.ReadToEnd());

                tr.Close();


            }


            File.WriteAllText(outputDestinationPath, sb.ToString());


        }
        public void Executenestnowbackfillrocessing(string strBSECSVArr, string datetostore, string name, int count, string mappingsymbol)
        {


            FileHelperEngine engineBSECSV1 = new FileHelperEngine(typeof(nestnow));





                    //Get BSE Equity Filename day, month, year
                string[] words = strBSECSVArr.Split('\\');

                    string strbseequityfilename = words[words.Length - 1];


                    nestnow[] resbsecsv1 = engineBSECSV1.ReadFile(strBSECSVArr) as nestnow[];


                    nestnowfinal [] finalarr = new nestnowfinal [resbsecsv1.Length];
                    int icntr = 0;
               

                    while (icntr < resbsecsv1.Length)
                    {
                        finalarr[icntr] = new nestnowfinal ();
                        
                        //finalarr[icntr].ticker = strbseequityfilename.Substring(0, strbseequityfilename.Length - 4);
                        //finalarr[icntr].name = strbseequityfilename.Substring(0, strbseequityfilename.Length - 4); ;

                        finalarr[icntr].ticker = resbsecsv1[icntr].Name ;
                        string[] datetime = resbsecsv1[icntr].datetime .Split(' ');
                        datetostore = datetime[0];
                        string timetostore = datetime[1];
                        finalarr[icntr].date = datetostore; // String.Format("{0:yyyyMMdd}", myDate);
                        finalarr[icntr].open = resbsecsv1[icntr].OPEN_PRICE;
                        finalarr[icntr].high = resbsecsv1[icntr].HIGH_PRICE;
                        finalarr[icntr].low = resbsecsv1[icntr].LOW_PRICE;
                        finalarr[icntr].close = resbsecsv1[icntr].CLOSE_PRICE;
                        finalarr[icntr].volume = resbsecsv1[icntr].volume;
                        finalarr[icntr].time = timetostore;



                        icntr++;
                    }

                    FileHelperEngine engineBSECSVFINAL = new FileHelperEngine(typeof(nestnowfinal ));

                    engineBSECSVFINAL.WriteFile(strBSECSVArr, finalarr);
                    log4net.Config.XmlConfigurator.Configure();
                    ILog log = LogManager.GetLogger(typeof(MainWindow));
                    log.Debug("Google File Processing ....... ");
                
                return;


            











        }
        public void ExecuteYAHOOProcessing(string[] strBSECSVArr, string datetostore, string name, int count,string mappingsymbol)
        {
         
            if (name == "GOOGLEEOD")
            {
                FileHelperEngine engineBSECSV1 = new FileHelperEngine(typeof(GOOGLE));

               


                foreach (string obj in strBSECSVArr)
                {

                    //Get BSE Equity Filename day, month, year
                    string[] words = obj.Split('\\');

                    string strbseequityfilename = words[words.Length - 1];


                    GOOGLE[] resbsecsv1 = engineBSECSV1.ReadFile(obj) as GOOGLE[];


                    GOOGLEFINAL[] finalarr = new GOOGLEFINAL[resbsecsv1.Length];
                    int icntr = 0;
                    //int hrs = Convert.ToInt32(GHRS.SelectedItem);
                    //int min = Convert.ToInt32(GMIN.SelectedItem);
                    //int hrstostore = Convert.ToInt32(hrs - 5);
                    //int mintostore = Convert.ToInt32(min - 30);
                    DateTime timefromyahoo = DateTime.Today;

                    //if (hrs > 5 && min > 30)
                    //{
                    //    timefromyahoo = new DateTime(1970, 1, 1, hrstostore, mintostore, 0).AddSeconds(Convert.ToInt64(resbsecsv1[icntr].Name));
                    //}
                    //else if (hrs > 5 && min <= 30)
                    //{
                    //    timefromyahoo = new DateTime(1970, 1, 1, hrstostore, 0, 0).AddSeconds(Convert.ToInt64(resbsecsv1[icntr].Name));
                    //}
                    //else if (hrs < 5 && min > 30)
                    //{
                    //    timefromyahoo = new DateTime(1970, 1, 1, 0, mintostore, 0).AddSeconds(Convert.ToInt64(resbsecsv1[icntr].Name));
                    //}
                    //else
                    //{

                    // }
                    long valueforgoogletime = 1;

                    while (icntr < resbsecsv1.Length)
                    {
                        finalarr[icntr] = new GOOGLEFINAL();
                        if (resbsecsv1[icntr].Name.Contains('a'))
                        {
                         valueforgoogletime = Convert.ToInt64(resbsecsv1[icntr].Name.Substring(1, resbsecsv1[icntr].Name.Length - 1));
                        }

                        timefromyahoo = new DateTime(1970, 1, 1, 5, 30, 0).AddSeconds(valueforgoogletime);
                        int mindata = 0;
                        if (google_time_frame.SelectedItem == "5 min")
                        {
                            mindata = 300;
                        }
                        else if (google_time_frame.SelectedItem == "1 min")
                        {
                            mindata = 60;
                        }
                        
                        valueforgoogletime = valueforgoogletime + mindata ;

                        string timetostore = timefromyahoo.Hour .ToString()+":"+timefromyahoo.Minute.ToString()+":"+timefromyahoo.Millisecond.ToString();


                        string[] yahoodate = timefromyahoo.ToString().Split('-');

                        datetostore =yahoodate[0]+yahoodate[1] +yahoodate[2].Substring(0, 4)  ;
                        //finalarr[icntr].ticker = strbseequityfilename.Substring(0, strbseequityfilename.Length - 4);
                        //finalarr[icntr].name = strbseequityfilename.Substring(0, strbseequityfilename.Length - 4); ;

                        finalarr[icntr].ticker =mappingsymbol ;
                            finalarr[icntr].name = mappingsymbol;

                      if (Format_cb.SelectedItem == "Metastock")
                        {

                          if(mappingsymbol.Contains("-"))
                          {
                              string[] symbolformeta = mappingsymbol.Split('-');
                              finalarr[icntr].ticker = symbolformeta[0];

                          }

                            finalarr[icntr].name = "I";
                            datetostore =yahoodate[2].Substring(2, 2)+ yahoodate[1] + yahoodate[0] ;

                        }
                        finalarr[icntr].date = datetostore; // String.Format("{0:yyyyMMdd}", myDate);
                        finalarr[icntr].open = resbsecsv1[icntr].OPEN_PRICE;
                        finalarr[icntr].high = resbsecsv1[icntr].HIGH_PRICE;
                        finalarr[icntr].low = resbsecsv1[icntr].LOW_PRICE;
                        finalarr[icntr].close = resbsecsv1[icntr].CLOSE_PRICE;
                        finalarr[icntr].volume = resbsecsv1[icntr].volume;
                        finalarr[icntr].time = timetostore ;

                        finalarr[icntr].openint = 0;  //enint;


                        icntr++;
                    }

                    FileHelperEngine engineBSECSVFINAL = new FileHelperEngine(typeof(GOOGLEFINAL));
                    if (Format_cb.SelectedItem != "Metastock")
                    {
                        engineBSECSVFINAL.HeaderText = "Ticker,Name,Date,Time,Open,High,Low,Close,Volume,OPENINT";
                    }


                    engineBSECSVFINAL.WriteFile(obj, finalarr);
                    log4net.Config.XmlConfigurator.Configure();
                    ILog log = LogManager.GetLogger(typeof(MainWindow));
                    log.Debug("Google File Processing ....... ");
                }
                return;


            }





            

           
            


        }
        public string yahootime(DateTime timetostore)
        {


            if (timetostore.Hour == 03)
            {
                if (timetostore.Minute > 30)
                {
                    return "19:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
                else
                {
                    return "20:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
            }
            if (timetostore.Hour == 04)
            {
                if (timetostore.Minute > 30)
                {
                    return "20:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
                else
                {
                    return "21:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
            }
            if (timetostore.Hour == 05)
            {
                if (timetostore.Minute > 30)
                {
                    return "21:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
                else
                {
                    return "22:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
            }
            if (timetostore.Hour == 06)
            {
                if (timetostore.Minute > 30)
                {
                    return "22:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
                else
                {
                    return "23:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
            }
            if (timetostore.Hour == 07)
            {
                if (timetostore.Minute > 30)
                {
                    return "23:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
                else
                {
                    return "24:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
            }
            if (timetostore.Hour == 08)
            {
                if (timetostore.Minute > 30)
                {
                    return "24:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
                else
                {
                    return "24:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
            }

            if (timetostore.Hour == 13)
            {
                if (timetostore.Minute > 30)
                {
                    return "00:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
                else
                {
                    return "00:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
            }



            else if (timetostore.Hour == 14)
            {
                if (timetostore.Minute < 30)
                {
                    return "00:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
                else
                {
                    return "01:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
            }
            else if (timetostore.Hour == 15)
            {
                if (timetostore.Minute < 30)
                {
                    return "01:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
                else
                {
                    return "02:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
            }
            else if (timetostore.Hour == 16)
            {
                if (timetostore.Minute < 30)
                {
                    return "02:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
                else
                {
                    return "03:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
            }

            else if (timetostore.Hour == 17)
            {
                if (timetostore.Minute < 30)
                {
                    return "03:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
                else
                {
                    return "04:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
            }
            else if (timetostore.Hour == 18)
            {
                if (timetostore.Minute < 30)
                {
                    return "04:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
                else
                {
                    return "05:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
            }

            else if (timetostore.Hour == 19)
            {
                if (timetostore.Minute < 30)
                {
                    return "05:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
                else
                {
                    return "06:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
            }
            else if (timetostore.Hour == 20)
            {
                if (timetostore.Minute < 30)
                {
                    return "06:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
                else
                {
                    return "07:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
            }
            else if (timetostore.Hour == 21)
            {
                if (timetostore.Minute < 30)
                {
                    return "07:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
                else
                {
                    return "08:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
            }
            else if (timetostore.Hour == 22)
            {
                if (timetostore.Minute < 30)
                {
                    return "08:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
                else
                {
                    return "09:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
            }
            else if (timetostore.Hour == 23)
            {
                if (timetostore.Minute < 30)
                {
                    return "09:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
                else
                {
                    return "10:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
            }
            else if (timetostore.Hour == 24)
            {
                if (timetostore.Minute < 30)
                {
                    return "10:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
                else
                {
                    return "11:" + timetostore.Minute.ToString() + ":" + timetostore.Second.ToString();

                }
            }


            return null;
        }
        private void downliaddata(string path, string url)
        {


            try
            {
                //If Data is Not Present For Date Then  Exception Occure And It Get Added Into List Box  
                // Client.DownloadFile("http://www.mcx-sx.com/downloads/daily/EquityDownloads/Market%20Statistics%20Report_" + date1 + ".csv.", File_path);

                log4net.Config.XmlConfigurator.Configure();
                ILog log = LogManager.GetLogger(typeof(MainWindow));
                log.Debug(url + "Download Started at " + DateTime.Now.ToString("HH:mm:ss tt"));

                Client.Headers.Add("Accept", "application/zip");
                Client.Headers.Add("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8");
                Client.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.1 (KHTML, like Gecko) Chrome/21.0.1180.89 Safari/537.1");
                Client.DownloadFile(url, path);


                log.Debug(url + "Download Completed at " + DateTime.Now.ToString("HH:mm:ss tt"));

                //string clientHeader = "DATE" + "," + "TICKER" + " " + "," + "NAME" + "," + " " + "," + " " + "," + "OPEN" + "," + "HIGH" + "," + "LOW" + "," + "CLOSE" + "," + "VOLUME" + "," + "OPENINT" + Environment.NewLine;

                //Format_Header(File_path, clientHeader);
            }
            catch (Exception ex)
            {
                if ((ex.ToString().Contains("404")) || (ex.ToString().Contains("400")))
                {
                    System.Windows.MessageBox.Show("File not downloaded please check symbol name ");
                    log4net.Config.XmlConfigurator.Configure();
                    ILog log = LogManager.GetLogger(typeof(MainWindow));
                    log.Warn("Data Not Found For " + url);

                }
            }


        }

        private void close_btn_Click(object sender, RoutedEventArgs e)
        {
            config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings.Remove("amipath");

            config.AppSettings.Settings.Add("amipath", db_path.Text);
            config.Save(ConfigurationSaveMode.Full);
            ConfigurationManager.RefreshSection("appSettings");
            config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings.Remove("targetpathforcombo");

            config.AppSettings.Settings.Add("targetpathforcombo", txtTargetFolder.Text);
            config.Save(ConfigurationSaveMode.Full);
            ConfigurationManager.RefreshSection("appSettings");
            config.AppSettings.Settings.Remove("amipath");

            config.AppSettings.Settings.Add("amipath", db_path.Text);
            config.Save(ConfigurationSaveMode.Full);
            ConfigurationManager.RefreshSection("appSettings");


          Environment.Exit(0);
        }

        private void google_symbol_add_Click(object sender, RoutedEventArgs e)
        {
           
        }

        private void button1_Click_1(object sender, RoutedEventArgs e)
        {

            if (db_path.Text == "")
            {
                System.Windows.MessageBox.Show("Please enter database path ");
                db_path.Focus();
                return;
            }
            ExcelType = Type.GetTypeFromProgID("Broker.Application");
            ExcelInst = Activator.CreateInstance(ExcelType);
            ExcelType.InvokeMember("Visible", BindingFlags.SetProperty, null,
                      ExcelInst, new object[1] { true });
            ExcelType.InvokeMember("LoadDatabase", BindingFlags.InvokeMethod | BindingFlags.Public, null,
                 ExcelInst, new string[1] { db_path.Text  });


            string targetpath = ConfigurationManager.AppSettings["targetpathforcombo"];
            string[] sharekhanfilePaths = Directory.GetFiles(targetpath + "\\sharekhan", "*.csv", SearchOption.TopDirectoryOnly);
            string[] odinfilePaths = Directory.GetFiles(targetpath + "\\odin", "*.csv", SearchOption.TopDirectoryOnly);
            string[] nestnowfilePaths = Directory.GetFiles(targetpath + "\\nest-now", "*.csv", SearchOption.TopDirectoryOnly);
            string datetostore = "";
            
            for (int i = 0; i < nestnowfilePaths.Count(); i++)
            {
                try
                {
                    Executenestnowbackfillrocessing(nestnowfilePaths[i], datetostore, "GOOGLEEOD", i, targetpath + "\\nest-now\\" + nestnowfilePaths[i].ToString());
                }
                catch
                {
                }
            }
            for (int i = 0; i < nestnowfilePaths.Count(); i++)
            {

                args[0] = Convert.ToInt16(0);
                args[1] = nestnowfilePaths[i];
                args[2] = "shubhanest-now.format";


                ExcelType.InvokeMember("Import", BindingFlags.InvokeMethod | BindingFlags.Public, null,
                          ExcelInst, args);
            }
            for (int i = 0; i < sharekhanfilePaths.Count(); i++)
            {
                args[0] = Convert.ToInt16(0);
                args[1] = sharekhanfilePaths[i];
                args[2] = "Shubhasharekhan.format";


                ExcelType.InvokeMember("Import", BindingFlags.InvokeMethod | BindingFlags.Public, null,
                          ExcelInst, args);
            }
            for (int i = 0; i < odinfilePaths.Count(); i++)
            {
                args[0] = Convert.ToInt16(0);
                args[1] = odinfilePaths[i];
                args[2] = "Shubhasharekhan.format";


                ExcelType.InvokeMember("Import", BindingFlags.InvokeMethod | BindingFlags.Public, null,
                          ExcelInst, args);
            }
        }

        

        private void remove_symbol_Click(object sender, RoutedEventArgs e)
        {
            string targetpath = ConfigurationManager.AppSettings["targetpathforcombo"];

            try
            {
                int indextoremove = list_rtsymbol.SelectedIndex;
                list_rtsymbol.Items.RemoveAt(indextoremove);
                list_google_symbol.Items.RemoveAt(indextoremove);
                mapping_symbol_list.Items.RemoveAt(indextoremove);

                System.IO.File.Delete(txtTargetFolder.Text + "\\NESTRt.txt");

                for (int i = 0; i < list_rtsymbol .Items.Count; i++)
                {
                    using (var writer = new StreamWriter(targetpath + "\\NESTRt.txt", true))
                        writer.WriteLine(list_rtsymbol.Items[i].ToString());
                }
                System.IO.File.Delete(targetpath + "\\shubha_google_symbols.txt");

                for (int i = 0; i < list_google_symbol .Items.Count; i++)
                {
                    using (var writer = new StreamWriter(targetpath + "\\shubha_google_symbols.txt", true))
                        writer.WriteLine(list_google_symbol.Items[i].ToString());
                }


                System.IO.File.Delete(targetpath + "\\shubha_mapping_symbol.txt");

                for (int i = 0; i < mapping_symbol_list.Items.Count; i++)
                {
                    using (var writer = new StreamWriter(targetpath + "\\shubha_mapping_symbol.txt", true))
                        writer.WriteLine(mapping_symbol_list.Items[i].ToString());
                }

            }
            catch
            {
            }
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings.Remove("amipath");

            config.AppSettings.Settings.Add("amipath", db_path.Text);
            config.Save(ConfigurationSaveMode.Full);
            ConfigurationManager.RefreshSection("appSettings");
            config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings.Remove("targetpathforcombo");

            config.AppSettings.Settings.Add("targetpathforcombo", txtTargetFolder.Text);
            config.Save(ConfigurationSaveMode.Full);
            ConfigurationManager.RefreshSection("appSettings");
            config.AppSettings.Settings.Remove("amipath");

            config.AppSettings.Settings.Add("amipath", db_path.Text);
            config.Save(ConfigurationSaveMode.Full);
            ConfigurationManager.RefreshSection("appSettings");

            config.AppSettings.Settings.Remove("terminalname");

            config.AppSettings.Settings.Add("terminalname", RTD_server_name.SelectedItem.ToString());
            config.Save(ConfigurationSaveMode.Full);
            ConfigurationManager.RefreshSection("appSettings");


            Environment.Exit(0);
        }

        private void remove_all_symbol_Click(object sender, RoutedEventArgs e)
        {
            string targetpath = ConfigurationManager.AppSettings["targetpathforcombo"];

            try
            {
                list_rtsymbol.Items.Clear();
                list_google_symbol.Items.Clear();
                mapping_symbol_list .Items.Clear();


                System.IO.File.Delete(txtTargetFolder.Text + "\\NESTRt.txt");

                for (int i = 0; i < list_rtsymbol.Items.Count; i++)
                {
                    using (var writer = new StreamWriter(targetpath + "\\NESTRt.txt", true))
                        writer.WriteLine(list_rtsymbol.Items[i].ToString());
                }
                System.IO.File.Delete(targetpath + "\\shubha_google_symbols.txt");

                for (int i = 0; i < list_google_symbol.Items.Count; i++)
                {
                    using (var writer = new StreamWriter(targetpath + "\\shubha_google_symbols.txt", true))
                        writer.WriteLine(list_google_symbol.Items[i].ToString());
                }
                System.IO.File.Delete(targetpath + "\\shubha_mapping_symbol.txt");

                for (int i = 0; i < mapping_symbol_list.Items.Count; i++)
                {
                    using (var writer = new StreamWriter(targetpath + "\\shubha_mapping_symbol.txt", true))
                        writer.WriteLine(mapping_symbol_list.Items[i].ToString());
                }

                

            }
            catch
            {
            }
        }

       
       

       
       
       

        
    }
}
