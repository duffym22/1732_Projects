using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace _NET_1732_Attendance
{
    /// <summary>
    /// Interaction logic for UserPrompt.xaml
    /// </summary>
    public partial class UserPrompt : Window
    {
        public UserPrompt()
        {
            InitializeComponent();
        }

        public string FirstName { get { return txt_FirstName.Text; } }
        public string LastName { get { return txt_LastName.Text; } }

        private void Btn_SubmitName_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }
    }
}
