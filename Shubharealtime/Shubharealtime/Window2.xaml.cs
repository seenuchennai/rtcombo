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
using System.Threading.Tasks;


namespace Shubharealtime
{
    /// <summary>
    /// Interaction logic for Window2.xaml
    /// </summary>
    public partial class Window2 : Window
    {
        public Window2()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Shubharealtime.datadownload s = new datadownload();
            Task.Factory.StartNew(s.serverintitilization);

        }

        private void EndRT_Click(object sender, RoutedEventArgs e)
        {

            Shubharealtime.datadownload s = new datadownload();
            s.stopdata();
            this.Close();
            Shubharealtime.Window1 w = new Window1();
            w.Close();

           
            
           
            string filepath = System.Reflection.Assembly.GetExecutingAssembly().Location.ToString();
            System.Diagnostics.Process.Start(filepath);
            Application.Current.Shutdown();

        }
    }
}
