using log4net;
using log4net.Config;
using System;
using System.Reflection;
using System.Text;
using System.Windows;

namespace _1732_Attendance
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        internal long ID_Scan;
        private GSheetsAPI gAPI;
        private static readonly ILog _log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        const string _DARKORANGE = "#FFFF8C00";
        const string _GREEN = "#FF008000";
        const string _RED = "#FFFF0000";
        const string _BLACK = "#FFFFFFFF";

        #region *** MAIN FORM ***
        public MainWindow()
        {
            InitializeComponent();
            XmlConfigurator.Configure();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Initialize();
        }

        private void btn_Login_Click(object sender, RoutedEventArgs e)
        {

        }

        private void btn_Reconnect_Click(object sender, RoutedEventArgs e)
        {
            Initialize();
        }

        private void txt_ID_Scan_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            try
            {
                if (e.Key == System.Windows.Input.Key.Enter || e.Key == System.Windows.Input.Key.Return)
                {
                    if (long.TryParse(txt_ID_Scan.Text, out ID_Scan))
                    {
                        if (Lookup_ID())
                        {
                            Update_Record();
                            txt_ID_Scan.Clear();
                        }
                        else
                        {
                            DisplayText(string.Format("{0} not registered. Please find a mentor to register your ID", ID_Scan.ToString()));
                            txt_ID_Scan.Clear();
                        }
                    }
                    else
                    {
                        Log(string.Format("Invalid ID Entry: {0}", ID_Scan.ToString()));
                        DisplayText("Invalid ID entered");
                        txt_ID_Scan.Clear();
                    }
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
                UI_Control(false); //if an exception is thrown - disable the UI
            }
            finally
            {
                if (txt_ID_Scan.IsVisible)
                {
                    txt_ID_Scan.Focus();
                }
            }
        }
        #endregion

        #region *** ACCESS METHODS ***

        private void Initialize()
        {
            gAPI = new GSheetsAPI();
            if (gAPI.AuthorizeGoogleApp())
            {
                Log("Application authorized");
                Log("Refreshing local data");
                if (gAPI.Refresh_Local_Data())
                {
                    Log("Successfully refreshed local data");
                    UI_Control(true);
                    txt_ID_Scan.Focus();
                }
                else
                {
                    Log("Failed to refresh local data");
                    UI_Control(false);
                    btn_Reconnect.Focus();
                }
            }
            else
            {
                Log("Unable to connect to Google Sheets");
                Log("Verify internet connectivity");
                Log("Verify API key still valid");
                Log(gAPI.LastException);
                UI_Control(false);
            }
        }

        private void UI_Control(bool state)
        {
            /// State == TRUE
            /// - Disable & Hide Reconnect button
            /// - Enable & Show ID text field entry
            /// State == FALSE
            /// - Enable & Show Reconnect button
            /// - Disable & Hide ID text field entry

            if (state)
            {
                Log("Enabling text entry field");
                txt_ID_Scan.IsEnabled = true;
                txt_ID_Scan.Visibility = Visibility.Visible;
                Log("Disabling reconnect button");
                btn_Reconnect.Visibility = Visibility.Hidden;
                btn_Reconnect.IsEnabled = false;

            }
            else
            {
                Log("Disabling text entry field");
                txt_ID_Scan.IsEnabled = false;
                txt_ID_Scan.Visibility = Visibility.Hidden;
                Log("Enabling reconnect button");
                btn_Reconnect.Visibility = Visibility.Visible;
                btn_Reconnect.IsEnabled = true;
            }
        }

        private bool Lookup_ID()
        {
            bool
                success = false;

            try
            {
                Log(string.Format("Verifying ID: {0}", ID_Scan.ToString()));
                if (gAPI.Check_Valid_ID(ID_Scan))
                {
                    Log(string.Format("ID: {0} - Verified", ID_Scan.ToString()));
                    success = true;
                }
                else
                {
                    Log(string.Format("ID: {0} - Not registered", ID_Scan.ToString()));
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return success;
        }

        private bool Update_Record()
        {
            bool
                success = false;
            try
            {
                Log(string.Format("Updating ID: {0}", ID_Scan.ToString()));
                if (gAPI.Update_User_Status(ID_Scan))
                {
                    DisplayText(string.Format("{0} - {1} - CHECKED {2}", ID_Scan.ToString(), gAPI.Get_ID_Name(ID_Scan), gAPI.Check_ID_Status(ID_Scan)));
                    Log(string.Format("ID: {0} | NAME: {1} | STATUS: {2}", ID_Scan.ToString(), gAPI.Get_ID_Name(ID_Scan), gAPI.Check_ID_Status(ID_Scan)));
                    success = true;
                }
                else
                {
                    Log(gAPI.LastException);
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return success;
        }

        internal void Log(string text)
        {
            _log.Info(text);
        }

        #endregion

        #region *** DISPLAY/EXCEPTION HANDLING ***
        internal async void DisplayText(string text)
        {
            //TextRange tr = new TextRange(rtb_Output.Document.ContentEnd, rtb_Output.Document.ContentEnd);
            //tr.Text = string.Format("{0}\r", text);
            //tr.ApplyPropertyValue(TextElement.ForegroundProperty, new BrushConverter().ConvertFromString(_BLACK));
            txt_Status.Text = text;
            await System.Threading.Tasks.Task.Delay(2500);
            txt_Status.Text = string.Empty;
        }

        internal void HandleException(Exception ex, string callingMethod)
        {
            StringBuilder _exMsg = new StringBuilder();

            _exMsg.AppendLine(string.Format("Exception thrown in: {0}", callingMethod));
            _exMsg.AppendLine(string.IsNullOrEmpty(ex.Message) ? "" : ex.Message);
            _exMsg.AppendLine(string.IsNullOrEmpty(ex.Source) ? "" : ex.Source);
            _exMsg.AppendLine(string.IsNullOrEmpty(ex.StackTrace.ToString()) ? "" : ex.StackTrace.ToString());

            Log(_exMsg.ToString());
        }
        #endregion
    }
}
