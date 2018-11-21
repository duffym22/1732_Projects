using log4net;
using log4net.Config;
using System.Drawing;
using System.Reflection;
using System.Windows;

namespace _1732_Attendance
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        internal string ID_Scan;
        private GSheetsAPI gAPI;
        private static readonly ILog _log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        public MainWindow()
        {
            InitializeComponent();
            XmlConfigurator.Configure();
        }

        private void btn_Login_Click(object sender, RoutedEventArgs e)
        {
            
        }

        private void Lookup_ID()
        {

        }

        private void Update_Record()
        {

        }

        private void txt_ID_Scan_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter || e.Key == System.Windows.Input.Key.Return)
            {
                ID_Scan = txt_ID_Scan.Text;
            }

        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            gAPI = new GSheetsAPI();
            if(!gAPI.AuthorizeGoogleApp())
            {
                Log("Unable to connect to Google Sheets");
                Log("Verify internet connectivity");
                Log("Verify API key still valid");
                Log(gAPI.LastException);
                //Disable UI
                //Enable and show "reconnect button" to try to reconnect to sheets
            }
            else
            {
                gAPI.Refresh_Local_Data();
            }
        }

        internal void Log(string text)
        {
            _log.Info(text);
        }

        internal void DisplayText(string text)
        {
            rtb_Output.AppendText(text);
        }

        internal void DisplayText(string text, Color color)
        {
            rtb_Output.AppendText(text);

        }
    }
}
