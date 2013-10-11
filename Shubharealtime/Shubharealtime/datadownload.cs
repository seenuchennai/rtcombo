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
using System.Threading.Tasks;
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
using System.Globalization;

namespace Shubharealtime
{
    class datadownload
    {
        Configuration config;

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
        
        public  void RtdataRecall()
        {

            string interval = ConfigurationManager.AppSettings["interval"];

            DispatcherTimer1.Tick += new EventHandler(dispatcherTimerForRT_Tick);
            DispatcherTimer1.Interval = new TimeSpan(0, 0, Convert.ToInt32(interval));
            DispatcherTimer1.Start();
            CommandManager.InvalidateRequerySuggested();

        }
        public void rtddata()
        {
          
            string targetpath = ConfigurationManager.AppSettings["targetpathforcombo"];
            string formatname = ConfigurationManager.AppSettings["format"];
            CommandManager.InvalidateRequerySuggested();
            try
            {
                if (File.Exists(targetpath + "\\realtimemetastock.csv"))
                {
                    File.Delete(targetpath + "\\realtimemetastock.csv");
                }
                if (!Directory.Exists(targetpath + "\\Fchart"))
                {
                    Directory.CreateDirectory(targetpath + "\\Fchart");
                }
                if (File.Exists(targetpath + "\\realtimefchart.csv"))
                {
                    File.Copy(targetpath + "\\realtimefchart.csv", targetpath + "\\Fchart\\Finalrealtimefchart.csv", true);
                    File.Delete(targetpath + "\\realtimefchart.csv");
                }
                if (File.Exists(targetpath + "\\YahooRealTimeData.csv"))
                {

                    File.Delete(targetpath + "\\YahooRealTimeData.csv");
                }
                int countformappingsymbol=0;
                List<string> mapsymbol = new List<string>();
                using (var reader1 = new StreamReader("C:\\myshubhalabha\\shubha_mapping_symbol.txt"))
                {
                    string line1 = null;
                    while ((line1 = reader1.ReadLine()) != null)
                    {
                        mapsymbol.Add(line1);

                    }
                }
                yahoortdata.Clear();
                int flagfortotaldatacount = 0;
                using (var reader = new StreamReader(targetpath + "\\NESTRt.txt"))
                {
                    string line = null;
                    int RTtopiccount = 0;
                    yahoortdata.Clear();

                    while ((line = reader.ReadLine()) != null)
                    {
                        CommandManager.InvalidateRequerySuggested();

                        yahoortname.Add(line);
                        Array retval;


                        int j = m_server.Heartbeat();

                        bool bolGetNewValue = true;
                        object[] arrayForSymbol = new object[2];

                        // RTtopiccount++;    //imp it change topic id 
                        CommandManager.InvalidateRequerySuggested();

                        arrayForSymbol[0] = line;
                        arrayForSymbol[1] = "Trading Symbol";


                        Array sysArrParams = (Array)arrayForSymbol;
                        m_server.ConnectData(RTtopiccount, sysArrParams, bolGetNewValue);

                        RTtopiccount++;    //imp it change topic id 
                        object[] arrayForLTT = new object[2];

                        CommandManager.InvalidateRequerySuggested();

                        arrayForLTT[0] = line;
                        arrayForLTT[1] = "LUT";

                        Array sysArrParams1 = (Array)arrayForLTT;
                        m_server.ConnectData(RTtopiccount, sysArrParams1, bolGetNewValue);

                        RTtopiccount++;    //imp it change topic id 

                        object[] arrayForLTP = new object[2];


                        arrayForLTP[0] = line;
                        arrayForLTP[1] = "LTP";

                        Array sysArrParams2 = (Array)arrayForLTP;
                        m_server.ConnectData(RTtopiccount, sysArrParams2, bolGetNewValue);

                        CommandManager.InvalidateRequerySuggested();

                        RTtopiccount++;    //imp it change topic id 

                        CommandManager.InvalidateRequerySuggested();


                        object[] arrayForVolume = new object[2];

                        arrayForVolume[0] = line;
                        arrayForVolume[1] = "Volume Traded Today";

                        Array sysArrParams3 = (Array)arrayForVolume;
                        m_server.ConnectData(RTtopiccount, sysArrParams3, bolGetNewValue);

                        CommandManager.InvalidateRequerySuggested();

                        RTtopiccount++;    //imp it change topic id 
                        object[] arrayForopenint = new object[2];

                        arrayForopenint[0] = line;
                        arrayForopenint[1] = "Open Interest";

                        Array sysArrParams4 = (Array)arrayForopenint;
                        m_server.ConnectData(RTtopiccount, sysArrParams4, bolGetNewValue);


                        RTtopiccount++;    //imp it change topic id 

                        CommandManager.InvalidateRequerySuggested();

                        retval = m_server.RefreshData(10);


                        for (int count = 0; count <= 4; count++)
                        {
                            m_server.DisconnectData(count);
                        }
                        foreach (var item in retval)
                        {

                            yahoortdata.Add(item.ToString());
                            CommandManager.InvalidateRequerySuggested();

                        }

                        m_server.ServerTerminate();
                        flagfortotaldatacount++;
                        CommandManager.InvalidateRequerySuggested();


                    }
                    CommandManager.InvalidateRequerySuggested();

                    string tempfilepath = targetpath + "\\RealTimeData.txt";
                    //log4net.Config.XmlConfigurator.Configure();
                    //ILog log = LogManager.GetLogger(typeof(MainWindow));
                    //log.Debug("Data Capturing At" + DateTime.Now.TimeOfDay);
                    string storeinfile1 = "";
                    CommandManager.InvalidateRequerySuggested();

                    //c=c+2 we not want 1st 3rd 5th and so on values.
                    int value = 5;
                    int flagtocheckfirstvaluefordate = 0;

                    for (int j = 5; j < yahoortdata.Count - 1; j = j + 10)
                    {
                        CommandManager.InvalidateRequerySuggested();

                        int c;
                        value = j + 5;
                        if (flagtocheckfirstvaluefordate == 0)
                        {
                            storeinfile1 =  storeinfile1;
                            flagtocheckfirstvaluefordate = 1;
                            CommandManager.InvalidateRequerySuggested();

                        }
                        else
                        {
                            storeinfile1 = storeinfile1 + " " ;

                            flagtocheckfirstvaluefordate = 1;

                        }
                        int flagmap = 0;
                        
                        for (c = j; c <= value - 1; c = c + 1)
                        {
                            CommandManager.InvalidateRequerySuggested();

                            if (flagmap == 0)
                            {
                                storeinfile1 = storeinfile1 + " " + mapsymbol[countformappingsymbol];
                                countformappingsymbol++;
                                flagmap++;

                            }
                            else
                            {

                                storeinfile1 = storeinfile1 + " " + yahoortdata[c].ToString();
                            }

                        }


                        CommandManager.InvalidateRequerySuggested();

                        storeinfile1 = storeinfile1 + "\r\n";


                    }

                    //if (storeinfile1.Contains(","))
                    //{
                    //    storeinfile1.Replace(",", ".");
                    //}


                    //if count is greater than data required then dont write it in file
                    if (yahoortdata.Count <= flagfortotaldatacount * 10)
                    {
                       

                        //<TICKER>,<NAME>,<PER>,<DATE>,<TIME>,<OPEN>,<HIGH>,<LOW>,<CLOSE>,<VOL>,<OPENINT>
                        using (var writer = new StreamWriter(tempfilepath))

                            writer.WriteLine(storeinfile1);




                      
                        if (formatname== "Amibroker")
                        {

                            //string realtimemetastock = "";

                            //string datastoreforami = "";
                            //int count = 0;
                            //using (var reader1 = new StreamReader(tempfilepath))
                            //{
                            //    string line1 = null;
                            //    while ((line1 = reader1.ReadLine()) != null)
                            //    {

                            //        string[] words = line1.Split(' ');
                            //        if (line1 != "")
                            //        {
                            //            if (count == 0)
                            //            {
                            //                realtimemetastock = realtimemetastock + "," + words[0] + "," + words[1] + "," + words[2] + "," + words[3] + "," + words[4] + "," + words[5];
                            //                count++;
                            //            }
                            //            else
                            //            {
                            //                realtimemetastock = realtimemetastock + "," + words[1] + "," + words[2] + "," + words[3] + "," + words[4] + "," + words[5] + "," + words[6];


                            //            }
                            //            realtimemetastock = realtimemetastock + "\r\n";
                            //        }
                            //    }
                            //}


                            string filename = targetpath + "\\AmibrokerRTdata.txt";
                            //   System.Windows.MessageBox.Show(realtimemetastock);
                            using (var writer = new StreamWriter(filename))
                                writer.WriteLine(storeinfile1);

                        }



                        if (formatname== "Fchart")
                        {
                            int count = 0;
                            using (var reader1 = new StreamReader(tempfilepath))
                            {
                                string line1 = null;
                                string realtimemetastock = "";
                                while ((line1 = reader1.ReadLine()) != null)
                                {

                                    string[] words = line1.Split(' ');
                                    if (line1 != "")
                                    {
                                        //if (count == 0)
                                        //{
                                        //    realtimemetastock = words[1] + "," + words[1] + "," + "I" + "," + words[0] + "," + words[2] + "," + words[3] + "," + words[3] + "," + words[3] + "," + words[3] + "," + words[4] + "," + words[5];
                                        //    count++;
                                        //}
                                        //else
                                        //{
                                        //    realtimemetastock = words[2] + "," + words[2] + "," + "I" + "," + words[1] + "," + words[3] + "," + words[4] + "," + words[4] + "," + words[4] + "," + words[4] + "," + words[5] + "," + words[6];

                                        //}
                                        //realtimemetastock = words[1] + "," + words[1] + "," + "I" + "," + words[2] + "," + words[3] + "," + words[4] + "," + words[4] + "," + words[4] + "," + words[4] + "," + words[5];
                                        if (words[1] == "")
                                        {
                                            string[] timeforfchart = words[4].Split(':');
                                            string timetostore = timeforfchart[0] + ":" + timeforfchart[1];
                                            realtimemetastock = words[2] + "," + words[3] + "," + timetostore + "," + words[5] + "," + words[5] + "," + words[5] + "," + words[5] + "," + words[6];
                                        }
                                        else
                                        {
                                            string[] timeforfchart = words[3].Split(':');
                                            string timetostore = timeforfchart[0] + ":" + timeforfchart[1];
                                            realtimemetastock = words[1] + "," + words[2] + "," + timetostore + "," + words[4] + "," + words[4] + "," + words[4] + "," + words[4] + "," + words[5];

                                        }
                                        string filename = targetpath + "\\realtimefchart.csv";
                                        //   System.Windows.MessageBox.Show(realtimemetastock);
                                        using (var writer = new StreamWriter(filename, true))
                                            writer.WriteLine(realtimemetastock);
                                    }
                                }
                            }

                        }

                        if (formatname== "Metastock")
                        {
                            int count = 0;
                            string filename = "";
                            using (var reader1 = new StreamReader(tempfilepath))
                            {
                                string line1 = null;
                                string realtimemetastock = "";
                                while ((line1 = reader1.ReadLine()) != null)
                                {

                                    if (line1 != "")
                                    {
                                        string[] words = line1.Split(' ');

                                        //if (count == 0)
                                        //{
                                        //    if (words[1].Contains('-'))
                                        //    {
                                        //        string[] ticker =words[1].Split('-');
                                        //        words[1] = ticker[0];

                                        //    }
                                        if (words[1]=="")
                                           {
                                            realtimemetastock = words[2] + "," + words[2] + "," + "I" + "," + words[3] + "," + words[4] + "," + words[5] + "," + words[5] + "," + words[5] + "," + words[5] + "," + words[6];
                                           }
                                           else
                                           {
                                               realtimemetastock = words[1] + "," + words[1] + "," + "I" + "," + words[2] + "," + words[3] + "," + words[4] + "," + words[4] + "," + words[4] + "," + words[4] + "," + words[5];

                                           }
                                          //  count++;
                                            
                                        //}
                                        //else
                                        //{
                                        //    if (words[2].Contains('-'))
                                        //    {
                                        //        string[] ticker = words[1].Split('-');
                                        //        words[2] = ticker[0];

                                        //    }
                                        //    realtimemetastock = words[2] + "," + words[2] + "," + "I" + "," + words[1] + "," + words[3] + "," + words[4] + "," + words[4] + "," + words[4] + "," + words[4] + "," + words[5] + "," + words[6];

                                        //}
                                        realtimemetastock = "<TICKER>,<NAME>,<PER>,<DATE>,<TIME>,<OPEN>,<HIGH>,<LOW>,<CLOSE>,<VOLUME>\r\n" + realtimemetastock;

                                        filename = targetpath + "\\realtimemetastock.csv";
                                        //   System.Windows.MessageBox.Show(realtimemetastock);
                                        using (var writer = new StreamWriter(filename, true))
                                            writer.WriteLine(realtimemetastock);
                                    }
                                }
                            }
                            if (!Directory.Exists(targetpath + "\\Intraday\\Metastock"))
                            {
                                Directory.CreateDirectory(targetpath + "\\Intraday\\Metastock");
                            }
                            // commandpromptcall(filename, targetpath + "\\Intraday\\Metastock\\realtimemetastock");
                            try
                            {

                                string filepath = System.Reflection.Assembly.GetExecutingAssembly().Location.ToString();
                                string processtostart = filepath.Substring(0, filepath.Length - 18) + "asc2ms.exe";

                                File.Copy(processtostart, targetpath + "\\asc2ms.exe", true);
                            }
                            catch
                            {
                            }
                            System.Diagnostics.Process process = new System.Diagnostics.Process();
                            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
                            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                            startInfo.FileName = "cmd.exe";
                            //startInfo.Arguments = "/C  C:\\asc2ms.exe -f C:\\data\\Metastock\\M.csv -r r -o C:\\data\\Metastock\\google\\e";
                            startInfo.Arguments = "/C  " + targetpath + "\\asc2ms.exe -f " + filename + " -r r -o " + targetpath + "\\Intraday\\Metastock\\realtimemetastock";
                            // startInfo.Arguments = @"/C  C:\asc2ms.exe -f C:\Documents and Settings\maheshwar\My Documents\BSe\Downloads\Googleeod -r r -o C:\Documents and Settings\maheshwar\My Documents\BSe\Downloads\Googleeod\Metastock\a" ;



                            process.StartInfo = startInfo;
                            process.Start();

                        }
                    }

                    CommandManager.InvalidateRequerySuggested();



                    if (formatname== "Amibroker")
                    {
                        ExcelType = Type.GetTypeFromProgID("Broker.Application");
                        ExcelInst = Activator.CreateInstance(ExcelType);
                        args[0] = Convert.ToInt16(0);
                        args[1] = targetpath + "\\RealTimeData.txt";
                        args[2] = "ShubhaRt.format";
                        ExcelType.InvokeMember("Import", BindingFlags.InvokeMethod | BindingFlags.Public, null,
                             ExcelInst, args);


                        CommandManager.InvalidateRequerySuggested();

                        ExcelType.InvokeMember("RefreshAll", BindingFlags.InvokeMethod | BindingFlags.Public, null,
                               ExcelInst, new object[1] { "" });
                    }
                }
            }
            catch(Exception ex)
            {

            }
        }
        public void nowdata()
        {
            string targetpath = ConfigurationManager.AppSettings["targetpathforcombo"];
            string formatname = ConfigurationManager.AppSettings["format"];
            CommandManager.InvalidateRequerySuggested();
            try
            {
                if (File.Exists(targetpath + "\\realtimemetastock.csv"))
                {
                    File.Delete(targetpath + "\\realtimemetastock.csv");
                }
                if (!Directory.Exists(targetpath + "\\Fchart"))
                {
                    Directory.CreateDirectory(targetpath + "\\Fchart");
                }
                if (File.Exists(targetpath + "\\realtimefchart.csv"))
                {
                    File.Copy(targetpath + "\\realtimefchart.csv", targetpath + "\\Fchart\\Finalrealtimefchart.csv", true);
                    File.Delete(targetpath + "\\realtimefchart.csv");
                }
                if (File.Exists(targetpath + "\\YahooRealTimeData.csv"))
                {
                    File.Delete(targetpath + "\\YahooRealTimeData.csv");
                }
                yahoortdata.Clear();
                int flagfortotaldatacount = 0;
                int countformappingsymbol = 0;
                List<string> mapsymbol = new List<string>();
                using (var reader1 = new StreamReader("C:\\myshubhalabha\\shubha_mapping_symbol.txt"))
                {
                    string line1 = null;
                    while ((line1 = reader1.ReadLine()) != null)
                    {
                        mapsymbol.Add(line1);

                    }
                }
                using (var reader = new StreamReader(targetpath + "\\NESTRt.txt"))
                {
                    string line = null;
                    int RTtopiccount = 0;
                    yahoortdata.Clear();

                    while ((line = reader.ReadLine()) != null)
                    {
                        CommandManager.InvalidateRequerySuggested();

                        yahoortname.Add(line);
                        Array retval;


                        int j = m_server.Heartbeat();

                        bool bolGetNewValue = true;
                        object[] arrayForSymbol = new object[3];

                        // RTtopiccount++;    //imp it change topic id 
                        CommandManager.InvalidateRequerySuggested();

                        arrayForSymbol[0] = "MktWatch";

                        arrayForSymbol[1] = line;
                        arrayForSymbol[2] = "Trading Symbol";


                        Array sysArrParams = (Array)arrayForSymbol;
                        m_server.ConnectData(RTtopiccount, sysArrParams, bolGetNewValue);

                        RTtopiccount++;    //imp it change topic id 
                        object[] arrayForLTT = new object[3];

                        CommandManager.InvalidateRequerySuggested();

                        arrayForLTT[0] = "MktWatch";

                        arrayForLTT[1] = line;
                        arrayForLTT[2] = "Last Trade Time";

                        Array sysArrParams1 = (Array)arrayForLTT;
                        m_server.ConnectData(RTtopiccount, sysArrParams1, bolGetNewValue);

                        RTtopiccount++;    //imp it change topic id 

                        object[] arrayForLTP = new object[3];

                        arrayForLTP[0] = "MktWatch";

                        arrayForLTP[1] = line;
                        arrayForLTP[2] = "Last Traded Price";

                        Array sysArrParams2 = (Array)arrayForLTP;
                        m_server.ConnectData(RTtopiccount, sysArrParams2, bolGetNewValue);

                        CommandManager.InvalidateRequerySuggested();

                        RTtopiccount++;    //imp it change topic id 

                        CommandManager.InvalidateRequerySuggested();


                        object[] arrayForVolume = new object[3];
                        arrayForVolume[0] = "MktWatch";

                        arrayForVolume[1] = line;
                        arrayForVolume[2] = "Volume Traded Today";

                        Array sysArrParams3 = (Array)arrayForVolume;
                        m_server.ConnectData(RTtopiccount, sysArrParams3, bolGetNewValue);

                        CommandManager.InvalidateRequerySuggested();

                        RTtopiccount++;    //imp it change topic id 
                        object[] arrayForopenint = new object[3];
                        arrayForopenint[0] = "MktWatch";

                        arrayForopenint[1] = line;
                        arrayForopenint[2] = "Open Interest";

                        Array sysArrParams4 = (Array)arrayForopenint;
                        m_server.ConnectData(RTtopiccount, sysArrParams4, bolGetNewValue);


                        RTtopiccount++;    //imp it change topic id 

                        CommandManager.InvalidateRequerySuggested();

                        retval = m_server.RefreshData(10);


                        for (int count = 0; count <= 4; count++)
                        {
                            m_server.DisconnectData(count);
                        }
                        foreach (var item in retval)
                        {

                            yahoortdata.Add(item.ToString());
                            CommandManager.InvalidateRequerySuggested();

                        }

                        m_server.ServerTerminate();
                        flagfortotaldatacount++;
                        CommandManager.InvalidateRequerySuggested();


                    }
                    CommandManager.InvalidateRequerySuggested();

                    string tempfilepath = targetpath + "\\RealTimeData.txt";
                    //log4net.Config.XmlConfigurator.Configure();
                    //ILog log = LogManager.GetLogger(typeof(MainWindow));
                    //log.Debug("Data Capturing At" + DateTime.Now.TimeOfDay);
                    string storeinfile1 = "";
                    CommandManager.InvalidateRequerySuggested();

                    //c=c+2 we not want 1st 3rd 5th and so on values.
                    int value = 5;
                    int flagtocheckfirstvaluefordate = 0;

                    for (int j = 5; j < yahoortdata.Count - 1; j = j + 10)
                    {
                        CommandManager.InvalidateRequerySuggested();

                        int c;
                        value = j + 5;
                        if (flagtocheckfirstvaluefordate == 0)
                        {
                            storeinfile1 = DateTime.Today.Date.ToShortDateString() + storeinfile1;
                            flagtocheckfirstvaluefordate = 1;
                            CommandManager.InvalidateRequerySuggested();

                        }
                        else
                        {
                            storeinfile1 = storeinfile1 + "," + DateTime.Today.Date.ToShortDateString();

                            flagtocheckfirstvaluefordate = 1;

                        }
                        int flagmap = 0;
                        for (c = j; c <= value - 1; c = c + 1)
                        {

                            if (flagmap == 0)
                            {
                                storeinfile1 = storeinfile1 + " " + mapsymbol[countformappingsymbol];
                                countformappingsymbol++;
                                flagmap++;

                            }
                            else
                            {

                                storeinfile1 = storeinfile1 + " " + yahoortdata[c].ToString();
                            }

                        }


                        CommandManager.InvalidateRequerySuggested();


                        //////////////////////////////////////








                        storeinfile1 = storeinfile1 + "\r\n";


                    }



                    //if count is greater than data required then dont write it in file
                    if (yahoortdata.Count <= flagfortotaldatacount * 10)
                    {
                        //<TICKER>,<NAME>,<PER>,<DATE>,<TIME>,<OPEN>,<HIGH>,<LOW>,<CLOSE>,<VOL>,<OPENINT>
                        using (var writer = new StreamWriter(tempfilepath))

                            writer.WriteLine(storeinfile1);

                        if (formatname== "Amibroker")
                        {
                            
                            using (var reader1 = new StreamReader(tempfilepath))
                            {
                                int count = 0;

                                string line1 = null;
                                string realtimemetastock = "";
                                while ((line1 = reader1.ReadLine()) != null)
                                {
                                    if(line1!="")
                                    {
                                    string[] words = line1.Split(' ');
                                    //if (count == 0)
                                    //{
                                    //    realtimemetastock = words[1] + words[0] + " " + words[2] + " " + words[3] +" " + words[4] + " " + words[5];
                                    //    count++;
                                    //}
                                    //else
                                    //{
                                    //    realtimemetastock = words[2] + words[1] + " " + words[3] + " " + words[4] + " " + words[5] + " " + words[6];

                                    //}

                                    if (words[0].Contains(","))
                                    {
                                        words[0] = words[0].Substring(1, words[0].Length - 1);
                                    }
                                   // realtimemetastock = words[1] + " " + words[0] + " " + words[2] + " " + words[3] + " " + words[4] + " " + words[5];
                                    realtimemetastock = words[1] + " " + words[0] + " " + words[2] + " " + words[3] + " " + words[4] + " " + words[5];
                                   
                                    
                                    string filename = targetpath + "\\YahooRealTimeData.csv";
                                    //   System.Windows.MessageBox.Show(realtimemetastock);
                                    using (var writer = new StreamWriter(filename, true))
                                        writer.WriteLine(realtimemetastock);
                                    }
                                }
                            }

                        }

                        if (formatname== "Fchart")
                        {
                            int count = 0;
                            using (var reader1 = new StreamReader(tempfilepath))
                            {
                                string line1 = null;
                                string realtimemetastock = "";
                                while ((line1 = reader1.ReadLine()) != null)
                                {

                                    string[] words = line1.Split(',');
                                    if (count == 0)
                                    {
                                        string[] timeforfchart = words[2].Split(':');
                                        string timetostore = timeforfchart[0] + ":" + timeforfchart[1];
                                        realtimemetastock = words[1] + "," + words[0] + "," + timetostore + "," + words[3] + "," + words[3] + "," + words[3] + "," + words[3] + "," + words[4] + "," + words[5];
                                        count++;
                                    }
                                    else
                                    {
                                        string[] timeforfchart = words[3].Split(':');
                                        string timetostore = timeforfchart[0] + ":" + timeforfchart[1];
                                        realtimemetastock = words[2] + "," + words[1] + "," + timetostore + "," + words[4] + "," + words[4] + "," + words[4] + "," + words[4] + "," + words[5] + "," + words[6];

                                    }
                                    string filename = targetpath + "\\realtimefchart.csv";
                                    //   System.Windows.MessageBox.Show(realtimemetastock);
                                    using (var writer = new StreamWriter(filename, true))
                                        writer.WriteLine(realtimemetastock);
                                }
                            }

                        }

                        if (formatname== "Metastock")
                        {
                            int count = 0;
                            string filename = "";
                            using (var reader1 = new StreamReader(tempfilepath))
                            {
                                string line1 = null;
                                string realtimemetastock = "";
                                while ((line1 = reader1.ReadLine()) != null)
                                {

                                    if (line1 != "")
                                    {
                                        string[] words = line1.Split(',');

                                        if (count == 0)
                                        {
                                            if (words[1].Contains('-'))
                                            {
                                                string[] ticker = words[1].Split('-');
                                                words[1] = ticker[0];

                                            }
                                            realtimemetastock = words[1] + "," + words[1] + "," + "I" + "," + words[0] + "," + words[2] + "," + words[3] + "," + words[3] + "," + words[3] + "," + words[3] + "," + words[4] + "," + words[5];
                                            count++;
                                        }
                                        else
                                        {
                                            if (words[2].Contains('-'))
                                            {
                                                string[] ticker = words[1].Split('-');
                                                words[2] = ticker[0];

                                            }
                                            realtimemetastock = words[2] + "," + words[2] + "," + "I" + "," + words[1] + "," + words[3] + "," + words[4] + "," + words[4] + "," + words[4] + "," + words[4] + "," + words[5] + "," + words[6];

                                        }
                                        realtimemetastock = "<TICKER>,<NAME>,<PER>,<DATE>,<TIME>,<OPEN>,<HIGH>,<LOW>,<CLOSE>,<VOLUME>\r\n" + realtimemetastock;

                                        filename = targetpath + "\\realtimemetastock.csv";
                                        //   System.Windows.MessageBox.Show(realtimemetastock);
                                        using (var writer = new StreamWriter(filename, true))
                                            writer.WriteLine(realtimemetastock);
                                    }
                                }
                            }
                            if (!Directory.Exists(targetpath + "\\Intraday\\Metastock"))
                            {
                                Directory.CreateDirectory(targetpath + "\\Intraday\\Metastock");
                            }
                            // commandpromptcall(filename, targetpath + "\\Intraday\\Metastock\\realtimemetastock");
                            System.Diagnostics.Process process = new System.Diagnostics.Process();
                            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
                            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                            startInfo.FileName = "cmd.exe";
                            //startInfo.Arguments = "/C  C:\\asc2ms.exe -f C:\\data\\Metastock\\M.csv -r r -o C:\\data\\Metastock\\google\\e";
                            startInfo.Arguments = "/C  C:\\asc2ms.exe -f " + filename + " -r r -o " + targetpath + "\\Intraday\\Metastock\\realtimemetastock";
                            // startInfo.Arguments = @"/C  C:\asc2ms.exe -f C:\Documents and Settings\maheshwar\My Documents\BSe\Downloads\Googleeod -r r -o C:\Documents and Settings\maheshwar\My Documents\BSe\Downloads\Googleeod\Metastock\a" ;



                            process.StartInfo = startInfo;
                            process.Start();

                        }


                    }

                    CommandManager.InvalidateRequerySuggested();



                    if (formatname== "Amibroker")
                    {


                        ExcelType = Type.GetTypeFromProgID("Broker.Application");
                        ExcelInst = Activator.CreateInstance(ExcelType);
                        args[0] = Convert.ToInt16(0);
                        args[1] = targetpath + "\\YahooRealTimeData.csv";
                        args[2] = "ShubhaRt.format";
                        ExcelType.InvokeMember("Import", BindingFlags.InvokeMethod | BindingFlags.Public, null,
                             ExcelInst, args);


                        CommandManager.InvalidateRequerySuggested();

                        ExcelType.InvokeMember("RefreshAll", BindingFlags.InvokeMethod | BindingFlags.Public, null,
                               ExcelInst, new object[1] { "" });
                    }
                }
            }
            catch
            {
                //CommandManager.InvalidateRequerySuggested();

                //log4net.Config.XmlConfigurator.Configure();
                //ILog log = LogManager.GetLogger(typeof(MainWindow));
                //log.Debug("Error While Data Capture ....");

                //CommandManager.InvalidateRequerySuggested();

            }
        }

        public void stopdata()
        {
            DispatcherTimer1.Stop();
        }
        private void dispatcherTimerForRT_Tick(object sender, EventArgs e)
        {

            string terminal = ConfigurationManager.AppSettings["terminal"];

            CommandManager.InvalidateRequerySuggested();

            if (terminal == "NEST")
            {

                rtddata();

            }
            else if (terminal == "NOW")
            {
                nowdata();

            }
            RtdataRecall();

        }
        public void serverintitilization()
        {
            string terminal = ConfigurationManager.AppSettings["terminal"];

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
            RtdataRecall();
        }

    }
}
