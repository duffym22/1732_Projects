using System.Windows;

namespace _1732_Attendance
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        internal string ID_Scan;

        public MainWindow()
        {
            InitializeComponent();
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

            /// 1. Connect to Google Sheet (team account)
            /// 2. Lookup student ID against database (worksheet)
            /// 3. Update student ID timestamp
            /// 

            /// ID | Name (LN, FN) | Timestamp | Status
            /// ex: 60007 | Duffy, Matthew | 2018-10-18 19:27:39 | IN
            /// ex: 60007 | Duffy, Matthew | 2018-10-18 20:27:39 | OUT
        }
    }
}
