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

        private void txt_scanID_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if(e.Key == System.Windows.Input.Key.Enter || e.Key == System.Windows.Input.Key.Return)
            {
                ID_Scan = txt_ID_Scan.Text;

            }
        }

        private void Lookup_ID()
        {

        }

        private void Update_Record()
        {

        }
    }
}
