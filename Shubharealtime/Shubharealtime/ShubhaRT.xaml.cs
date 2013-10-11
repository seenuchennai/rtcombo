using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Navigation;
using System.Configuration;
using Microsoft.Win32;
using System.Security.Principal ;
using System.Diagnostics;
using System.Reflection;

namespace Shubharealtime
{
    /// <summary>
    /// Interaction logic for ShubhaRT.xaml
    /// </summary>
    public partial class ShubhaRT : Window
    {
        public ShubhaRT()
        {
            InitializeComponent();
        }
        private bool IsRunAsAdministrator()
        {
            var wi = WindowsIdentity.GetCurrent();
            var wp = new WindowsPrincipal(wi);

            return wp.IsInRole(WindowsBuiltInRole.Administrator);
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            ////////////////////////////
           
            if (!IsRunAsAdministrator())
            {
                // It is not possible to launch a ClickOnce app as administrator directly, so instead we launch the
                // app as administrator in a new process.
                var processInfo = new ProcessStartInfo(Assembly.GetExecutingAssembly().CodeBase);

                // The following properties run the new process as administrator
                processInfo.UseShellExecute = true;
                processInfo.Verb = "runas";

                // Start the new process
                try
                {
                    Process.Start(processInfo);
                }
                catch (Exception)
                {
                    // The user did not allow the application to run as administrator
                    MessageBox.Show("Sorry, this application must be run as Administrator.");
                }

                // Shut down the current process
                Application.Current.Shutdown();
            }
            else
            {
                // We are running as administrator

                // Do normal startup stuff.

                try
                {
                    System.Net.WebRequest myRequest = System.Net.WebRequest.Create("http://www.Google.co.in");
                    System.Net.WebResponse myResponse = myRequest.GetResponse();
                    Uri a = new System.Uri("http://shubhalabha.in/eng/ads/www/delivery/afr.php?zoneid=22&amp;target=_blank&amp;cb=INSERT_RANDOM_NUMBER_HERE");
                  //  Uri a = new System.Uri(" http://shubhalabha.in/products-2/");

                    advertisement.Source = a;
                    //  wad4.Source = a4;


                }
                catch
                {


                    advertisement.Visibility = Visibility.Hidden;


                }
            }
        }
        void wb_LoadCompleted(object sender, NavigationEventArgs e)
        {
            string script = "document.body.style.overflow ='hidden'";
            System.Windows.Controls.WebBrowser wb = (System.Windows.Controls.WebBrowser)sender;
            wb.InvokeScript("execScript", new Object[] { script, "JavaScript" });
        }
        private void button1_Click(object sender, RoutedEventArgs e)
        {
            RegistryKey regKey = Registry.CurrentUser;
            regKey = regKey.CreateSubKey(@"Windows-temp\");

            try
            {
                var registerdate = regKey.GetValue("sd");
                var paidornot = regKey.GetValue("sp");

                var chktmp = regKey.GetValue("ApplicationId");

                //if user delete register entry then show login window agian 
                if (chktmp != null)
                {
                    if (chktmp.ToString() == "1" && registerdate != null && paidornot != null)
                    {
                        string date = DateTime.Today.Date.ToShortDateString();
                        if (!date.Contains("-"))
                        {
                            MessageBox.Show("your system date format is  "+date +" please set date format to DD-mm-yyyy");
                        }
                        this.Hide();
                        Shubharealtime.Window1 w = new Window1();
                        w.ShowDialog();
                    }
                    else
                    {
                        string date = DateTime.Today.Date.ToShortDateString();
                        if (!date.Contains("-"))
                        {
                            MessageBox.Show("your system date format is  " + date + " please set date format to DD-mm-yyyy");
                        }
                        this.Hide();
                        Shubharealtime.MainWindow w = new MainWindow();
                        w.ShowDialog();
                    }
                }
                else
                {
                    this.Hide();
                    Shubharealtime.MainWindow w = new MainWindow();
                    w.ShowDialog();
                }




            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message );
            }
            
        }
    }
}
