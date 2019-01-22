using log4net;
using log4net.Config;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;

namespace _NET_1732_Attendance
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        #region *** VARIABLES ***
        const string _LOGIN = "Login";
        const string _EXIT = "Exit";

        const string _REGULAR_MODE_SCAN = "Scan your ID";
        const string _MENTOR_MODE_SCAN = "Scan/Enter your Mentor ID";
        const string _LOG_FILE_PATH = @"C:\1732_Attendance\proper.log";

        private GSheetsAPI gAPI;

        private Timer
           timer;

        private DispatcherTimer
           displayTimer;

        private DateTime
           lastDisplayUpdate;

        private static readonly ILog
           _log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        #endregion

        #region *** PROPERTIES ***
        private bool Mentor_Mode { get; set; }
        private ulong Logged_In_Mentor_ID { get; set; }
        private string Log_File_Path { get; set; }
        #endregion

        #region *** MAIN FORM ***
        public MainWindow()
        {
            InitializeComponent();
            XmlConfigurator.Configure();

            displayTimer = new DispatcherTimer();
            displayTimer.Tick += DisplayTimer_Tick;
            displayTimer.Interval = new TimeSpan(0, 0, 5);
            displayTimer.Start();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Initialize();
        }

        private void DisplayTimer_Tick(object sender, EventArgs e)
        {
            if ((DateTime.Now - lastDisplayUpdate).TotalSeconds > 5)
            {
                txt_Status.Text = string.Empty;
            }
        }

        private void BTN_Login_Click(object sender, RoutedEventArgs e)
        {
            if (BTN_Login.Content.Equals(_LOGIN))
            {
                UI_Control(3); //Set UI to mentor mode
            }
            else if (BTN_Login.Content.Equals(_EXIT))
            {
                UI_Control(4); //Set login to normal scanning
                UI_Control(0); //Set UI to normal scanning
            }
        }

        private void BTN_Reconnect_Click(object sender, RoutedEventArgs e)
        {
            Initialize();
        }

        private void BTN_Check_In_User_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(TXT_Card_ID.Text))
                {
                    if (Parse_Scanned_ID(TXT_Card_ID.Text, out ulong ID))
                    {
                        if (Lookup_ID(ID, out ulong primaryID))
                        {
                            if (gAPI.Update_User_Status(primaryID))
                            {
                                DisplayAdminText(string.Format("User checked in - ID: {0}", primaryID));
                                Log(string.Format("User checked in - ID: {0}", primaryID));
                            }
                            else
                            {
                                DisplayAdminText(gAPI.LastException);
                                Log(gAPI.LastException);
                            }
                        }
                        else
                        {
                            DisplayAdminText(string.Format("ID - {0} is not registered.", ID.ToString()));
                        }
                    }
                    else
                    {
                        DisplayAdminText(string.Format("Invalid ID scanned. Please try a different card to check-in user", TXT_Card_ID.Text));
                        Log(string.Format("Invalid ID scanned to check-in user", TXT_Card_ID.Text));
                    }
                }
                else
                {
                    DisplayAdminText("Please scan/enter an ID of the user to check-in");
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            finally
            {
                TXT_Card_ID.Clear();
                TXT_Card_ID.Focus();
            }
        }

        private void BTN_Force_Checkout_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(TXT_Card_ID.Text))
                {
                    if (Parse_Scanned_ID(TXT_Card_ID.Text, out ulong ID))
                    {
                        if (Lookup_ID(ID, out ulong primaryID))
                        {
                            if (gAPI.Force_Logoff_User(primaryID))
                            {
                                DisplayAdminText(string.Format("User force checked out - ID: {0}", primaryID));
                                Log(string.Format("User force checked out - ID: {0}", primaryID));
                            }
                            else
                            {
                                DisplayAdminText(gAPI.LastException);
                                Log(gAPI.LastException);
                            }
                        }
                        else
                        {
                            DisplayAdminText(string.Format("ID - {0} is not registered.", ID.ToString()));
                        }
                    }
                    else
                    {
                        DisplayAdminText(string.Format("Invalid ID scanned. Please try a different card to force checkout user", TXT_Card_ID.Text));
                        Log(string.Format("Invalid ID scanned to force checkout user", TXT_Card_ID.Text));
                    }
                }
                else
                {
                    DisplayAdminText("Please scan/enter an ID of the user to force check-out");
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            finally
            {
                TXT_Card_ID.Clear();
                TXT_Card_ID.Focus();
            }
        }

        private void BTN_Add_User_Click(object sender, RoutedEventArgs e)
        {

            ulong
                secondaryID = 0;

            try
            {
                if (!string.IsNullOrEmpty(TXT_Card_ID.Text) && !string.IsNullOrEmpty(TXT_First_Name.Text) && !string.IsNullOrEmpty(TXT_Last_Name.Text))
                {
                    if (Parse_Scanned_ID(TXT_Card_ID.Text, out ulong ID))
                    {
                        if (Lookup_ID(ID, out ulong primaryID))
                        {
                            DisplayAdminText(string.Format("ID: {0} is already registered", primaryID.ToString()));
                            Log(string.Format("ID: {0} is already registered", primaryID.ToString()));
                        }
                        else
                        {
                            string fullName = string.Format("{0}, {1}", TXT_Last_Name.Text, TXT_First_Name.Text);

                            //if a secondary ID is specified, try to parse it and add it for the user
                            if (!string.IsNullOrEmpty(TXT_Printed_ID.Text))
                            {
                                if (Parse_Scanned_ID(TXT_Printed_ID.Text, out secondaryID))
                                {
                                    Log(string.Format("Successfully parsed secondary ID: {0}", secondaryID.ToString()));
                                }
                            }

                            if (gAPI.Add_User(ID, secondaryID, fullName, Logged_In_Mentor_ID, (bool)CHK_Is_Mentor.IsChecked))
                            {
                                DisplayAdminText(string.Format("Successfully added ID: {0} | NAME: {1}", ID.ToString(), fullName));
                                Log(string.Format("Mentor: {0} added ID: {1} | NAME: {2}", Logged_In_Mentor_ID.ToString(), ID.ToString(), fullName));
                            }
                            else
                            {
                                DisplayAdminText(string.Format("Failed to add ID: {0} | NAME: {1}", ID.ToString(), fullName));
                                Log(string.Format("Mentor: {0} failed to add ID: {1} | NAME: {2}", Logged_In_Mentor_ID.ToString(), ID.ToString(), fullName));
                                DisplayAdminText(gAPI.LastException);
                                Log(gAPI.LastException);
                            }
                        }
                    }
                    else
                    {
                        DisplayAdminText(string.Format("Invalid ID scanned. Please try a different card to register the user", TXT_Card_ID.Text));
                        Log(string.Format("Invalid ID scanned to add user", TXT_Card_ID.Text));
                    }

                    TXT_Card_ID.Clear();
                    TXT_Printed_ID.Clear();
                    TXT_Card_ID.Focus();
                    TXT_First_Name.Clear();
                    TXT_Last_Name.Clear();
                    CHK_Is_Mentor.IsChecked = false;
                }
                else
                {
                    DisplayAdminText("Please scan an ID and enter a name into the fields to add a user");
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        private void BTN_Update_User_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(TXT_Card_ID.Text))
                {
                    if (Parse_Scanned_ID(TXT_Card_ID.Text, out ulong ID))
                    {
                        if (Lookup_ID(ID, out ulong primaryID))
                        {
                            string[] name = gAPI.Get_ID_Name(primaryID).Split(',');
                            TXT_Last_Name.Text = name[0].Trim();
                            TXT_First_Name.Text = name[1].Trim();
                            CHK_Is_Mentor.IsChecked = gAPI.Check_Is_Mentor(primaryID);
                            UI_Display_Update_Options(true);
                            DisplayAdminText("User data imported. Please make any changes to the user by updating the fields and press Save to finish");
                        }
                        else
                        {
                            TXT_Card_ID.Clear();
                            DisplayAdminText(string.Format("ID - {0} is not registered.", ID.ToString()));
                        }
                    }
                    else
                    {
                        DisplayAdminText(string.Format("Invalid ID scanned. Please try a different card to update the user", TXT_Card_ID.Text));
                        Log(string.Format("Invalid ID scanned to update user", TXT_Card_ID.Text));
                    }

                }
                else
                {
                    DisplayAdminText("Please scan/enter an ID of the user to update");
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        private void BTN_Save_Updated_User_Click(object sender, RoutedEventArgs e)
        {
            ulong
                secondaryID = 0;

            try
            {
                if (!string.IsNullOrEmpty(TXT_First_Name.Text) && !string.IsNullOrEmpty(TXT_Last_Name.Text))
                {
                    if (Parse_Scanned_ID(TXT_Card_ID.Text, out ulong ID))
                    {
                        string fullName = string.Format("{0}, {1}", TXT_Last_Name.Text, TXT_First_Name.Text);

                        //if a secondary ID is specified, try to parse it and add it for the user
                        if (!string.IsNullOrEmpty(TXT_Printed_ID.Text))
                        {
                            if (Parse_Scanned_ID(TXT_Printed_ID.Text, out secondaryID))
                            {
                                Log(string.Format("Successfully parsed secondary ID: {0}", secondaryID.ToString()));
                            }
                        }

                        if (gAPI.Update_User(ID, secondaryID, fullName, Logged_In_Mentor_ID, (bool)CHK_Is_Mentor.IsChecked))
                        {
                            DisplayAdminText(string.Format("Successfully update ID: {0} | NAME: {1}", ID.ToString(), fullName));
                            Log(string.Format("Mentor: {0} updated ID: {1} | NAME: {2}", Logged_In_Mentor_ID.ToString(), ID.ToString(), fullName));
                        }
                        else
                        {
                            DisplayAdminText(string.Format("Failed to update ID: {0} | NAME: {1}", ID.ToString(), fullName));
                            Log(string.Format("Mentor: {0} failed to update ID: {1} | NAME: {2}", Logged_In_Mentor_ID.ToString(), ID.ToString(), fullName));
                            DisplayAdminText(gAPI.LastException);
                            Log(gAPI.LastException);
                        }
                    }
                    TXT_Card_ID.Clear();
                    TXT_Printed_ID.Clear();
                    TXT_First_Name.Clear();
                    TXT_Last_Name.Clear();
                    CHK_Is_Mentor.IsChecked = false;
                    TXT_Card_ID.Focus();

                    UI_Display_Update_Options(false);
                }
                else
                {
                    DisplayAdminText("Please enter the first and last name of the user to update");
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        private void BTN_Delete_User_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(TXT_Card_ID.Text))
                {
                    if (Parse_Scanned_ID(TXT_Card_ID.Text, out ulong ID))
                    {
                        if (Lookup_ID(ID, out ulong primaryID))
                        {
                            if (gAPI.Delete_User(primaryID, Logged_In_Mentor_ID))
                            {
                                DisplayAdminText(string.Format("Successfully deleted ID: {0}", primaryID.ToString()));
                                Log(string.Format("Mentor: {0} deleted ID: {1}", Logged_In_Mentor_ID.ToString(), primaryID.ToString()));
                            }
                            else
                            {
                                DisplayAdminText(string.Format("Failed to delete ID: {0}", primaryID.ToString()));
                                Log(string.Format("Mentor: {0} failed to delete ID: {1}", Logged_In_Mentor_ID.ToString(), primaryID.ToString()));
                                DisplayAdminText(gAPI.LastException);
                                Log(gAPI.LastException);
                            }
                        }
                        else
                        {
                            DisplayAdminText(string.Format("ID - {0} is not registered.", ID.ToString()));
                        }
                    }
                    else
                    {
                        DisplayAdminText(string.Format("Invalid ID scanned. Please try a different card to delete the user", TXT_Card_ID.Text));
                        Log(string.Format("Invalid ID scanned to delete user", TXT_Card_ID.Text));
                    }
                }
                else
                {
                    DisplayAdminText("Please scan/enter an ID to delete into the ID field");
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            finally
            {
                TXT_Card_ID.Clear();
                TXT_Card_ID.Focus();
            }
        }

        private void BTN_Who_Checked_In_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //clear datagrid
                List<User> users = new List<User>();
                UserDataGrid.ItemsSource = users;

                users = gAPI.Get_CheckedIn_Users();
                if (users.Count > 0)
                {
                    UserDataGrid.Visibility = Visibility.Visible;
                    DisplayAdminText(string.Format("Displayed currently logged in users. Count = {0}", users.Count));
                    Log(string.Format("Displayed currently logged in users. Count = {0}", users.Count));
                    UserDataGrid.ItemsSource = users;
                }
                else
                {
                    DisplayAdminText("No users currently logged in");
                    Log("No users currently logged in");
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        private void BTN_Recently_Checked_In_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //clear datagrid
                List<User> users = new List<User>();
                UserDataGrid.ItemsSource = users;

                users = gAPI.Get_Recently_CheckedIn_Users();
                if (users.Count > 0)
                {
                    UserDataGrid.Visibility = Visibility.Visible;
                    Log(string.Format("Displayed recently logged in users. Count = {0}", users.Count));
                    DisplayAdminText(string.Format("Displayed recently logged in users. Count = {0}", users.Count));
                    UserDataGrid.ItemsSource = users;
                }
                else
                {
                    DisplayAdminText("No users recently logged in");
                    Log("No users recently logged in");
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        private void BTN_Recently_Checked_Out_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //clear datagrid
                List<User> users = new List<User>();
                UserDataGrid.ItemsSource = users;

                users = gAPI.Get_Recently_CheckedOut_Users();
                if (users.Count > 0)
                {
                    UserDataGrid.Visibility = Visibility.Visible;
                    Log(string.Format("Displayed recently logged out users. Count = {0}", users.Count));
                    DisplayAdminText(string.Format("Displayed recently logged out users. Count = {0}", users.Count));
                    UserDataGrid.ItemsSource = users;
                }
                else
                {
                    DisplayAdminText("No users recently logged out");
                    Log("No users recently logged out");
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        private void BTN_Refresh_Main_Click(object sender, RoutedEventArgs e)
        {
            Refresh_Data();
        }

        private void BTN_Refresh_Click(object sender, RoutedEventArgs e)
        {
            Refresh_Data();
        }

        private void BTN_View_Log_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                List<string> log = gAPI.Get_Log_100_Rows();
                log.Reverse();
                if (log.Count > 0)
                {
                    DisplayAdminText("=== LOG ENTRIES ===");
                    foreach (string item in log)
                    {
                        DisplayAdminText(item);
                    }
                }
                else
                {
                    DisplayAdminText("No log entries");
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        private void BTN_Show_All_Users_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //clear datagrid
                List<User> users = new List<User>();
                UserDataGrid.ItemsSource = users;

                users = gAPI.Get_All_Users();
                if (users.Count > 0)
                {
                    UserDataGrid.Visibility = Visibility.Visible;
                    Log(string.Format("Displayed list of all users. Count = {0}", users.Count));
                    DisplayAdminText(string.Format("Displayed list of all users. Count = {0}", users.Count));
                    UserDataGrid.ItemsSource = users;
                }
                else
                {
                    Log("No users registered");
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        private void BTN_Add_Hours_Present_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(TXT_Card_ID.Text) && !string.IsNullOrEmpty(TXT_Hours.Text))
                {
                    if (Parse_Scanned_ID(TXT_Card_ID.Text, out ulong ID))
                    {
                        if (Lookup_ID(ID, out ulong primaryID))
                        {
                            if (gAPI.Credit_User_Hours(primaryID, TXT_Hours.Text, Logged_In_Mentor_ID))
                            {
                                DisplayAdminText(string.Format("Successfully credited {0} to ID: {1}", TXT_Hours.Text, primaryID.ToString()));
                                Log(string.Format("Successfully credited {0} to ID: {1}", TXT_Hours.Text, primaryID.ToString()));
                            }
                            else
                            {
                                DisplayAdminText(string.Format("Failed to credit {0} to ID: {1}", TXT_Hours.Text, primaryID.ToString()));
                                Log(string.Format("Mentor: {0} failed to credit ID: {1} with {2}", Logged_In_Mentor_ID.ToString(), primaryID.ToString(), TXT_Hours.Text));
                                DisplayAdminText(gAPI.LastException);
                                Log(gAPI.LastException);
                            }
                        }
                        else
                        {
                            TXT_Card_ID.Clear();
                            DisplayAdminText(string.Format("ID - {0} is not registered.", ID.ToString()));
                        }

                        TXT_Card_ID.Clear();
                        TXT_Hours.Clear();

                    }
                    else
                    {
                        DisplayAdminText(string.Format("Invalid ID scanned. Please try a different card to update the user", TXT_Card_ID.Text));
                        Log(string.Format("Invalid ID scanned to update user", TXT_Card_ID.Text));
                    }

                }
                else
                {
                    DisplayAdminText("Please scan/enter an ID of the user and enter the number of time present (hours, minutes and seconds) in the format 00:00:00");
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        private void BTN_Add_Hours_Missed_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(TXT_Card_ID.Text) && !string.IsNullOrEmpty(TXT_Hours.Text))
                {
                    if (Parse_Scanned_ID(TXT_Card_ID.Text, out ulong ID))
                    {
                        if (Lookup_ID(ID, out ulong primaryID))
                        {
                            if (gAPI.Add_Missed_Hours(primaryID, TXT_Hours.Text, Logged_In_Mentor_ID))
                            {
                                DisplayAdminText(string.Format("Successfully added missed hours ({0}) to ID: {1}", TXT_Hours.Text, primaryID.ToString()));
                                Log(string.Format("Successfully added missed hours ({0}) to ID: {1}", TXT_Hours.Text, primaryID.ToString()));
                            }
                            else
                            {
                                DisplayAdminText(string.Format("Failed to add missed hours ({0}) to ID: {1}", TXT_Hours.Text, primaryID.ToString()));
                                Log(string.Format("Mentor: {0} failed to add missed hours ({1}) to ID: {2}", Logged_In_Mentor_ID.ToString(), TXT_Hours.Text, primaryID.ToString()));
                                Log(gAPI.LastException);
                            }
                        }
                        else
                        {
                            TXT_Card_ID.Clear();
                            DisplayAdminText(string.Format("ID - {0} is not registered.", ID.ToString()));
                        }

                        TXT_Card_ID.Clear();
                        TXT_Hours.Clear();

                    }
                    else
                    {
                        DisplayAdminText(string.Format("Invalid ID scanned. Please try a different card to update the ID: {0}", TXT_Card_ID.Text));
                        Log(string.Format("Invalid ID scanned to add missed hours to ID: {0}", TXT_Card_ID.Text));
                    }
                }
                else
                {
                    DisplayAdminText("Please scan/enter an ID of the user and enter the number of missed time (hours, minutes and seconds) in the format 00:00:00");
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        private void TXT_ID_Scan_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            string
                idText = string.Empty;

            try
            {
                if (e.Key == System.Windows.Input.Key.Enter || e.Key == System.Windows.Input.Key.Return)
                {
                    idText = TXT_Scan.Text;
                    TXT_Scan.Clear();
                    if (Parse_Scanned_ID(idText, out ulong ID))
                    {
                        if (Lookup_ID(ID, out ulong primaryID))
                        {
                            //Mentor Admin mode
                            if (Mentor_Mode)
                            {
                                if (Verify_Mentor_ID(primaryID))
                                {
                                    UI_Control(2);
                                    DisplayAdminText(string.Format("Mentor Authorized - ID: {0} | NAME: {1}", primaryID, gAPI.Get_ID_Name(primaryID)));
                                }
                                else
                                {
                                    DisplayText(string.Format("ID - {0} is not authorized as a Mentor", primaryID.ToString()));
                                    Log(string.Format("ID - {0} is not authorized as a Mentor", primaryID.ToString()));
                                }
                            }
                            else
                            {
                                //Regular scanning mode
                                Log(string.Format("Updating ID: {0}", primaryID.ToString()));
                                Update_Record(primaryID);
                                //Thread t = new Thread(() => Update_Record(ID_Scan));
                                //t.Start();
                            }
                        }
                        else
                        {
                            Log(string.Format("ID - {0} is not registered", ID.ToString()));
                            DisplayText(string.Format("ID ({0}) is not registered. Please find a Mentor to register your ID", ID.ToString()));
                            Log_Unregistered_User(ID);
                        }
                    }
                    else
                    {
                        DisplayText("Invalid ID entered");
                        Log(string.Format("Invalid ID Entry: {0}", idText));
                    }
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
                UI_Control(1); //if an exception is thrown - disable the UI
            }
            finally
            {
                if (TXT_Scan.IsVisible)
                {
                    TXT_Scan.Focus();
                    TXT_Scan.InvalidateVisual();

                }
            }
        }

        private void Log_Unregistered_User(ulong ID)
        {
            if (!gAPI.Log_Unregistered_User(ID))
            {
                Log(gAPI.LastException);
            }
        }

        private void UserDataGrid_LostFocus(object sender, RoutedEventArgs e)
        {
            UserDataGrid.Visibility = Visibility.Hidden;
        }

        private void UserDataGrid_FocusableChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            ///idk something
            ///
        }

        private void Setup_Checkout_Timer()
        {
            TimeSpan timeToGo = gAPI.Auto_Checkout_Time - DateTime.Now;

            if (DateTime.Today.DayOfWeek.Equals(DayOfWeek.Saturday) || DateTime.Today.DayOfWeek.Equals(DayOfWeek.Sunday))
            {
                gAPI.Team_Checkout_Time = Parse_Checkout_Time(ConfigurationManager.AppSettings["WEEKEND_TEAM_CHECKOUT_TIME"]);
            }
            else
            {
                gAPI.Team_Checkout_Time = Parse_Checkout_Time(ConfigurationManager.AppSettings["WEEKDAY_TEAM_CHECKOUT_TIME"]);
            }

            if (timeToGo.Ticks < 0)
            {
                timeToGo = new TimeSpan(1, 0, 0, 0) - timeToGo.Negate();
            }
            Log(string.Format("Time until next auto-check out of users: {0}", timeToGo.ToString()));

            timer = new Timer(x =>
            {
                Check_Out_Users();
            }, null, timeToGo, Timeout.InfiniteTimeSpan);
        }

        #endregion

        #region *** ACCESS METHODS ***

        private void Initialize()
        {
            int
                sheetSelection = -1;

            string
                defaultLogo,
                customLogo,
                prodSheet,
                testSheet;

            gAPI = new GSheetsAPI
            {
                GID_Attendance_Status = Convert.ToInt32(ConfigurationManager.AppSettings["GID_ATTENDANCE_STATUS"]),
                GID_Accumulated_Hours = Convert.ToInt32(ConfigurationManager.AppSettings["GID_ACCUMULATED_HOURS"]),
                GID_Attendance_Log = Convert.ToInt32(ConfigurationManager.AppSettings["GID_ATTENDANCE_LOG"]),
                Recent_Time_Check = Convert.ToInt32(ConfigurationManager.AppSettings["RECENT_CHECKOUT_TIME"]),
                Auto_Checkout_Enabled = Convert.ToBoolean(ConfigurationManager.AppSettings["AUTO_CHECKOUT_ENABLED"])
            };

            ///Quick switcher in config to switch between prod and test sheets. 
            /// Value = 0 --> Use Prod sheet
            /// Value = 1 --> Use Test sheet
            sheetSelection = Convert.ToInt32(ConfigurationManager.AppSettings["SHEET_SELECTION"]);
            prodSheet = ConfigurationManager.AppSettings["PROD_SHEET_ID"];
            testSheet = ConfigurationManager.AppSettings["TEST_SHEET_ID"];
            gAPI.Sheet_ID = sheetSelection.Equals(0) ? prodSheet : testSheet;
            Log(string.Format("{0} sheet selected.", sheetSelection.Equals(0) ? "PROD" : "TEST"));

            if (DateTime.Today.DayOfWeek.Equals(DayOfWeek.Saturday) || DateTime.Today.DayOfWeek.Equals(DayOfWeek.Sunday))
            {
                gAPI.Team_Checkout_Time = Parse_Checkout_Time(ConfigurationManager.AppSettings["WEEKEND_TEAM_CHECKOUT_TIME"]);
            }
            else
            {
                gAPI.Team_Checkout_Time = Parse_Checkout_Time(ConfigurationManager.AppSettings["WEEKDAY_TEAM_CHECKOUT_TIME"]);
            }

            if (gAPI.Auto_Checkout_Enabled)
            {
                gAPI.Auto_Checkout_Time = Parse_Checkout_Time(ConfigurationManager.AppSettings["AUTO_CHECKOUT_TIME"]);
            }
            else
            {
                Log("Auto checkout disabled. Skipping parsing auto checkout time");
            }

            defaultLogo = ConfigurationManager.AppSettings["DEFAULT_LOGO"];
            customLogo = ConfigurationManager.AppSettings["CUSTOM_LOGO"];
            //if (!string.IsNullOrEmpty(customLogo))
            //{
            //    var path = Path.Combine(Environment.CurrentDirectory, "img", customLogo);
            //    var uri = new Uri(path);
            //    var bitmap = new BitmapImage(uri);
            //    IMG_Logo.Source = bitmap;
            //}
            //else
            //{
            //    var path = Path.Combine(Environment.CurrentDirectory, "img", defaultLogo);
            //    var uri = new Uri(path);
            //    var bitmap = new BitmapImage(uri);
            //    IMG_Logo.Source = bitmap;
            //}

            Log_File_Path = ConfigurationManager.AppSettings["LOG_FILE_PATH"];

            if (gAPI.AuthorizeGoogleApp())
            {
                Log("Application authorized");
                Log("Refreshing local data");
                if (gAPI.Refresh_Local_Data())
                {
                    Log("Successfully refreshed local data");
                    UI_Control(0);
                    if (gAPI.Auto_Checkout_Enabled)
                    {
                        Setup_Checkout_Timer();
                    }
                }
                else
                {
                    DisplayText("Failed to refresh local data");
                    Log("Failed to refresh local data");
                    UI_Control(1);
                    BTN_Reconnect.Focus();
                }
            }
            else
            {
                Log("Unable to connect to Google Sheets");
                Log("Verify internet connectivity");
                Log("Verify API key still valid");
                Log(gAPI.LastException);
                UI_Control(1);
            }
        }

        private DateTime Parse_Checkout_Time(string time)
        {
            DateTime value = new DateTime();
            try
            {
                Log("Auto checkout enabled. Attempting to parse auto checkout time");
                DateTime.TryParse(time, out value);
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return value;
        }

        private void UI_Control(int state)
        {
            /// State == TRUE
            /// - Disable & Hide Reconnect button
            /// - Enable & Show ID text field entry
            /// State == FALSE
            /// - Enable & Show Reconnect button
            /// - Disable & Hide ID text field entry

            switch (state)
            {
                //Normal Scanning 
                case 0:
                    Log("Enabling normal UI mode for scanning");
                    TXT_Scan.IsEnabled = true;
                    TXT_Scan.Visibility = Visibility.Visible;
                    Keyboard.Focus(TXT_Scan);


                    BTN_Refresh_Main.IsEnabled = true;
                    BTN_Refresh_Main.Visibility = Visibility.Visible;

                    GRD_Admin.IsEnabled = false;
                    GRD_Admin.Visibility = Visibility.Hidden;

                    BTN_Save_Updated_User.IsEnabled = false;
                    BTN_Save_Updated_User.Visibility = Visibility.Hidden;

                    BTN_Reconnect.Visibility = Visibility.Hidden;
                    BTN_Reconnect.IsEnabled = false;

                    UserDataGrid.Visibility = Visibility.Hidden;
                    break;
                //UI Disabled - Need to reconnect
                case 1:
                    Log("Disabling UI mode for scanning");
                    TXT_Scan.IsEnabled = true;
                    TXT_Scan.Visibility = Visibility.Visible;

                    BTN_Refresh_Main.IsEnabled = true;
                    BTN_Refresh_Main.Visibility = Visibility.Visible;

                    GRD_Admin.IsEnabled = false;
                    GRD_Admin.Visibility = Visibility.Hidden;

                    BTN_Save_Updated_User.IsEnabled = false;
                    BTN_Save_Updated_User.Visibility = Visibility.Hidden;

                    BTN_Reconnect.Visibility = Visibility.Visible;
                    BTN_Reconnect.IsEnabled = true;

                    UserDataGrid.Visibility = Visibility.Hidden;
                    break;
                //Mentor Mode
                case 2:
                    Log("Enabling mentor mode");
                    TXT_Scan.IsEnabled = false;
                    TXT_Scan.Visibility = Visibility.Hidden;

                    BTN_Refresh_Main.IsEnabled = false;
                    BTN_Refresh_Main.Visibility = Visibility.Hidden;

                    GRD_Admin.IsEnabled = true;
                    GRD_Admin.Visibility = Visibility.Visible;
                    Keyboard.Focus(TXT_Card_ID);

                    BTN_Save_Updated_User.IsEnabled = false;
                    BTN_Save_Updated_User.Visibility = Visibility.Hidden;

                    BTN_Reconnect.Visibility = Visibility.Hidden;
                    BTN_Reconnect.IsEnabled = false;

                    UserDataGrid.Visibility = Visibility.Hidden;
                    break;
                //Login - Setup Mentor Mode
                case 3:
                    Mentor_Mode = true;
                    LBL_ScanID.Text = _MENTOR_MODE_SCAN;
                    BTN_Login.Content = _EXIT;
                    TXT_Scan.Clear();
                    TXT_Scan.Focus();
                    break;
                //Login - Setup Normal Scanning
                case 4:
                    Mentor_Mode = false;
                    LBL_ScanID.Text = _REGULAR_MODE_SCAN;
                    BTN_Login.Content = _LOGIN;
                    Logged_In_Mentor_ID = 0;
                    RTB_AdminOutput.Document.Blocks.Clear();
                    break;
            }
        }

        private void UI_Display_Update_Options(bool state)
        {
            if (state)
            {
                BTN_Save_Updated_User.IsEnabled = true;
                BTN_Save_Updated_User.Visibility = Visibility.Visible;
            }
            else
            {
                BTN_Save_Updated_User.IsEnabled = false;
                BTN_Save_Updated_User.Visibility = Visibility.Hidden;
            }
        }

        private bool Parse_Scanned_ID(string scan, out ulong ID)
        {
            bool
                success = false;

            Regex
                alphanumeric = new Regex(@"^\w+$"),
                magStripeGeneric = new Regex(@"^%.*?$"),
                magStripeThreeTrack = new Regex(@"^%(B)?[\d]+\^[\w, \/]+\^[\d]+\?;[\d]+=[\d]+\?$"),
                numeric = new Regex(@"^\d+$");

            string
                temp,
                temp2;

            ID = 0;
            try
            {

                /// Otherwise, check to see if the input length is greater than 10
                /// - if not, then just check to see if it is numeric only and parse it
                ///  
                /// If the input length is greater than 10, check to see if it matches 
                /// any of the regex formats in this order
                /// 1. Numeric only - accept & truncate to first 10 digits
                /// 2. Generic mag stripe
                ///     2a. Check to see if it matches a full mag stripe (3 track) card swipe.
                ///         If so, take track one and check that it is only numbers. 
                ///         If so, truncate to first 10 digits
                ///     2b. If not a full stripe - check to see if it is a numeric single track, 
                ///         if so, truncate to first 10 digits
                /// 3. Check to see if the input is alphanumeric
                ///    If so, the input is rejected due to requirement that IDs are numeric only

                if (scan.Length > 10)
                {
                    //ID length is GREATER THAN 10
                    if (numeric.IsMatch(scan))
                    {
                        //ID is NUMERIC
                        temp = scan.Substring(0, 10);
                        DisplayAdminText(string.Format("Numeric ID is too long. Shortening to {0} characters", temp.Length));
                        ID = Parse_ID_Numeric(temp);
                        if (!ID.Equals(0))
                        {
                            success = true;
                        }
                    }
                    else if (magStripeGeneric.IsMatch(scan))
                    {
                        //ID matches a generic mag stripe
                        if (magStripeThreeTrack.IsMatch(scan))
                        {
                            DisplayAdminText(string.Format("3 track mag stripe detected. Parsing..."));
                            //ID matches a 3 track mag stripe
                            temp = scan.Substring(1, scan.Length - 1);              //remove Start (SS) and End (ES) Sentinels 
                            string[] contents = temp.Split('^');                    //split string on karet
                            temp2 = contents[0].ToUpper();                          //take first track and make all characters uppercase
                            temp2 = Regex.Replace(temp2, "[A-Za-z]", "");           //remove any instances of a letter
                            if (numeric.IsMatch(temp2))
                            {
                                temp2 = temp2.Substring(0, 10);
                                DisplayAdminText(string.Format("Numeric ID is too long. Shortening to {0} characters", temp2.Length));
                                ID = Parse_ID_Numeric(temp2);
                                if (!ID.Equals(0))
                                {
                                    success = true;
                                }
                            }
                            else
                            {
                                DisplayAdminText(string.Format("Cannot parse 3 track mag stripe: {0}", temp2));
                                Log(string.Format("Cannot parse 3 track mag stripe: {0}", temp2));
                            }
                        }
                        else
                        {
                            DisplayAdminText(string.Format("Generic mag stripe detected. Parsing..."));
                            temp = scan.Substring(1, scan.Length - 1);              //remove Start (SS) and End (ES) Sentinels 
                            temp = Regex.Replace(temp, @"[^\d]", string.Empty);     //remove any instances of a letter
                            DisplayAdminText(string.Format("Successfully parsed mag stripe to {0}.", temp));

                            //check to see if input is numeric 
                            //if not then check to see if it is alphanumeric for logging purposes 
                            if (numeric.IsMatch(temp))
                            {
                                //ID is NUMERIC
                                //in case the length remaining is GREATER THAN 10 - take a substring
                                if (temp.Length > 10)
                                {
                                    DisplayAdminText(string.Format("Numeric ID is too long. Shortening to {0} characters", temp.Length));
                                    temp = temp.Substring(0, 10);
                                }

                                ID = Parse_ID_Numeric(temp);
                                if (!ID.Equals(0))
                                {
                                    success = true;
                                }
                            }
                            else if (alphanumeric.IsMatch(temp))
                            {
                                //ID is ALPHANUMERIC
                                DisplayAdminText(string.Format("Cannot parse mag stripe alphanumeric ID: {0}", scan));
                                Log(string.Format("Cannot parse mag stripe alphanumeric ID: {0}", scan));
                            }
                            else
                            {
                                DisplayAdminText(string.Format("Unknown mag stripe input. Unable to parse input: {0}", scan));
                                Log(string.Format("Unknown mag stripe input. Unable to parse input: {0}", scan));
                            }
                        }
                    }
                    else if (alphanumeric.IsMatch(scan))
                    {
                        DisplayAdminText(string.Format("Cannot parse alphanumeric input: {0}", scan));
                        Log(string.Format("Cannot parse alphanumeric input: {0}", scan));
                    }
                }
                else
                {
                    //ID length is LESS THAN 10
                    if (numeric.IsMatch(scan))
                    {
                        //ID is NUMERIC
                        ID = Parse_ID_Numeric(scan);
                        if (!ID.Equals(0))
                        {
                            success = true;
                        }
                    }
                    else if (alphanumeric.IsMatch(scan))
                    {
                        //ID is ALPHANUMERIC
                        DisplayAdminText(string.Format("Cannot parse alphanumeric input: {0}", scan));
                        Log(string.Format("Cannot parse alphanumeric input: {0}", scan));
                    }
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return success;
        }

        private ulong Parse_ID_Numeric(string scan)
        {
            ulong ID = 0;
            try
            {
                if (ulong.TryParse(scan, out ID))
                {
                    DisplayAdminText(string.Format("Successfully parsed ID: {0}", scan));
                    Log(string.Format("Successfully parsed ID: {0}", scan));
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return ID;
        }

        private bool Lookup_ID(ulong ID, out ulong primaryID)
        {
            bool
                success = false;

            primaryID = 0;
            try
            {
                Log(string.Format("Verifying ID: {0}", ID.ToString()));
                if (gAPI.Check_Valid_ID(ID, out primaryID))
                {
                    if (!ID.Equals(primaryID))
                    {
                        Log(string.Format("Secondary ID: {0} - Verified", ID.ToString()));
                    }
                    else
                    {
                        Log(string.Format("ID: {0} - Verified", ID.ToString()));
                    }

                    success = true;
                }
                else
                {
                    Log(string.Format("ID: {0} - Not registered", ID.ToString()));
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return success;
        }

        private bool Verify_Mentor_ID(ulong ID)
        {
            bool
                success = false;

            try
            {
                Log(string.Format("Verifying ID is authorized: {0}", ID.ToString()));
                if (gAPI.Check_Is_Mentor(ID))
                {
                    Log(string.Format("ID: {0} - Authorized", ID.ToString()));
                    Logged_In_Mentor_ID = ID;
                    success = true;
                }
                else
                {
                    Log(string.Format("ID: {0} - Unauthorized", ID.ToString()));
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return success;
        }

        private void Update_Record(ulong ID)
        {
            try
            {
                if (gAPI.Update_User_Status(ID))
                {
                    DisplayText(string.Format("{0} - {1} - CHECKED {2}", ID.ToString(), gAPI.Get_ID_Name(ID), gAPI.Check_ID_Status(ID)));
                    Log(string.Format("ID: {0} | NAME: {1} | STATUS: {2}", ID.ToString(), gAPI.Get_ID_Name(ID), gAPI.Check_ID_Status(ID)));
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
        }

        private void Check_Out_Users()
        {
            try
            {
                //get list of users still checked in

                List<User> stillCheckedInUsers = gAPI.Get_CheckedIn_Users();
                Log(string.Format("{0} users remain checked in.", stillCheckedInUsers.Count));

                if (gAPI.Auto_Checkout_Enabled)
                {
                    foreach (User item in stillCheckedInUsers)
                    {
                        Log(string.Format("User forgot to check out - ID: {0} | NAME: {1}", item.ID.ToString(), item.Name.ToString()));
                        if (gAPI.Force_Logoff_User(item.ID))
                        {
                            Log(string.Format("User force checked out - ID: {0} | NAME: {1}", item.ID.ToString(), item.Name.ToString()));
                        }
                        else
                        {
                            Log(gAPI.LastException);
                        }
                    }
                    Log(string.Format("Force checked out {0} users", stillCheckedInUsers.Count));
                }
                else
                {
                    Log("Auto check-out disabled.");
                    Log(string.Format("{0} users remain checked in.", stillCheckedInUsers.Count));
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            finally
            {
                if (gAPI.Auto_Checkout_Enabled)
                {
                    Setup_Checkout_Timer();
                }
            }
        }

        internal void Log(string text)
        {
            _log.Info(text);
        }

        private void Refresh_Data()
        {
            try
            {
                if (gAPI.Refresh_Local_Data())
                {
                    Log("Local data refreshed");
                    if (Mentor_Mode)
                    {
                        DisplayAdminText("Local data refreshed");
                    }
                    else
                    {
                        DisplayText("Local data refreshed");
                    }
                }
                else
                {
                    Log("Failed to refresh local data");
                    Log(gAPI.LastException);
                    if (Mentor_Mode)
                    {
                        DisplayAdminText("Failed to refresh local data");
                    }
                    else
                    {
                        DisplayText("Failed to refresh local data");
                    }
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        #endregion

        #region *** DISPLAY/EXCEPTION HANDLING ***
        public delegate void Delegate_DisplayText(string text);

        internal void DisplayText(string text)
        {
            if (!txt_Status.Dispatcher.CheckAccess())
            {
                Dispatcher.BeginInvoke(new Delegate_DisplayText(DisplayText), text);
            }
            else
            {
                txt_Status.Text = text;
                lastDisplayUpdate = DateTime.Now;
            }
        }

        internal void DisplayAdminText(string text)
        {
            if (Mentor_Mode)
            {
                RTB_AdminOutput.AppendText(text + Environment.NewLine);
                RTB_AdminOutput.ScrollToEnd();
            }
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
