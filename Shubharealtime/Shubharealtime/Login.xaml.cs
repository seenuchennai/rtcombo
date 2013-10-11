using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Web;
using System.IO;
using System.Net;
using Microsoft.Win32;
using System.Configuration;

namespace Shubharealtime
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            string filepath = System.Reflection.Assembly.GetExecutingAssembly().Location.ToString();
            string leadurl = filepath.Substring(0, filepath.Length - 18) + "realtimereg.html";
            Uri a3 = new System.Uri("http://shubhalabha.in/real/realtimereg.html");
            lead.Source = a3 ;
            
        }

        private void Login_btn_Click(object sender, RoutedEventArgs e)
        {
              //user authentication
            try
            {
                string loginUri = "http://shubhalabha.in/community/wp-login.php";

                string reqString = "log=" + username.Text + "&pwd=" + password.Password ;
                byte[] requestData = Encoding.UTF8.GetBytes(reqString);

                CookieContainer cc = new CookieContainer();
                var request = (HttpWebRequest)WebRequest.Create(loginUri);
                request.Proxy = null;
                request.AllowAutoRedirect = false;
                request.CookieContainer = cc;
                request.Method = "post";

                request.ContentType = "application/x-www-form-urlencoded";
                request.ContentLength = requestData.Length;
                using (Stream s = request.GetRequestStream())
                    s.Write(requestData, 0, requestData.Length);

                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                {
                    int count = 1;
                    foreach (Cookie c in response.Cookies)
                    {
                        //responce 2 contain loggen in or not 
                        if (count == 2)
                        {
                            if (c.ToString().Contains("wordpress_logged_in_17e90d9fdb1ef2a442ed2d6aeb707f54"))
                            {
                                System.Windows.Forms.MessageBox.Show("Thank you for using shubhaRt plugin ");
                                try
                                {
                                   
                                    SetRegKey();
                                     this.Hide();
                                    Shubharealtime.Window1  s = new Window1();
                                    s.ShowDialog();
                                    return;

                                }
                                catch
                                {

                                }
                            }
                            else
                            {
                                System.Windows.Forms.MessageBox.Show("Please Enter Valid UserName & Password ");

                            }
                        }
                        else
                        {
                            count++;
                        }
                    }
                }
            }
            catch
            {

            }
            }
        

       
        public void SetRegKey()
        {

            //RegistryKey regKey = Registry.CurrentUser;
            //regKey = regKey.CreateSubKey(@"Software-RT\");
            //regKey.SetValue("ApplicationID", "1");

            RegistryKey regKey = Registry.CurrentUser;
            regKey = regKey.CreateSubKey(@"Windows-temp\");
            regKey.SetValue("sd", DateTime.Today.Date.ToString());
            regKey.SetValue("sp", "Key for xp");
            regKey.SetValue("ApplicationID", "1");
            regKey.SetValue("Applicationpath", System.Reflection.Assembly.GetExecutingAssembly().Location.ToString());


            
        }

        private void Regiser_btn_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://shubhalabha.in/community/wp-login.php?action=register");

        }

        private void Cancle_btn_Click(object sender, RoutedEventArgs e)
        {
            Environment.Exit(0);

        }

        private void lead_Loaded(object sender, RoutedEventArgs e)
        {
           
        }

        private void lead_Navigated(object sender, System.Windows.Navigation.NavigationEventArgs e)
        {
            if (lead.Source.ToString() == "http://shubhalabha.in/eng/crm/index.php?entryPoint=WebToLeadCapture")
            {
                SetRegKey();
                System.Windows.Forms.MessageBox.Show("Thank you for using shubha real time  ");


                this.Hide();
                Shubharealtime.Window1 s = new Window1();
                s.ShowDialog();

                return;
            }
           
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }

       

       
    }
}
