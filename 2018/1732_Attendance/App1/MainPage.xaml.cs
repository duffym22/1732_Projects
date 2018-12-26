using log4net;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Threading;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Documents;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace _1732_Attendance
{
   /// <summary>
   /// An empty page that can be used on its own or navigated to within a Frame.
   /// </summary>
   public sealed partial class MainPage : Page
   {
      #region *** VARIABLES ***
      const string _LOGIN = "Login";
      const string _EXIT = "Exit";

      const string _REGULAR_MODE_SCAN = "Scan your ID";
      const string _MENTOR_MODE_SCAN = "Scan/Enter your Mentor ID";

      private GSheetsAPI gAPI;
      private Timer timer;
      private static readonly ILog _log = LogManager.GetLogger(MethodBase.GetMethodFromHandle(new RuntimeMethodHandle()).DeclaringType);
      #endregion

      #region *** PROPERTIES ***
      private bool Mentor_Mode { get; set; }
      #endregion

      #region *** MAIN FORM ***
      public MainPage()
      {
         InitializeComponent();
      }

      private void BTN_Login_Click(object sender, RoutedEventArgs e)
      {

      }

      private void TXT_ID_Scan_KeyDown(object sender, KeyRoutedEventArgs e)
      {

      }

      private void BTN_Who_CheckedIn_Click(object sender, RoutedEventArgs e)
      {

      }

      private void BTN_Recently_Checked_In_Click(object sender, RoutedEventArgs e)
      {

      }

      private void BTN_Recently_Checked_Out_Click(object sender, RoutedEventArgs e)
      {

      }

      private void BTN_Refresh_Click(object sender, RoutedEventArgs e)
      {

      }

      private void BTN_Add_User_Click(object sender, RoutedEventArgs e)
      {

      }

      private void BTN_Update_User_Click(object sender, RoutedEventArgs e)
      {

      }

      private void BTN_Delete_User_Click(object sender, RoutedEventArgs e)
      {

      }

      private void BTN_Save_Updated_User_Click(object sender, RoutedEventArgs e)
      {
         try
         {
            ulong.TryParse(TXT_ID.Text, out ulong ID);
            if (!string.IsNullOrEmpty(TXT_First_Name.Text) && !string.IsNullOrEmpty(TXT_Last_Name.Text))
            {

               string fullName = string.Format("{0}, {1}", TXT_Last_Name.Text, TXT_First_Name.Text);
               if (gAPI.Update_User(ID, fullName, (bool)CHK_Is_Mentor.IsChecked))
               {
                  DisplayAdminText(string.Format("Successfully update ID: {0} | NAME: {1}", TXT_ID.Text, fullName));
                  Log(string.Format("Updated ID: {0} | NAME: {1}", TXT_ID.Text, fullName));
               }
               else
               {
                  DisplayAdminText(string.Format("Failed to update ID: {0} | NAME: {1}", TXT_ID.Text, fullName));
                  Log(string.Format("Failed to update ID: {0} | NAME: {1}", TXT_ID.Text, fullName));
                  Log(gAPI.LastException);
               }

               TXT_ID.Text = string.Empty;
               TXT_First_Name.Text = string.Empty;
               TXT_Last_Name.Text = string.Empty;
               CHK_Is_Mentor.IsChecked = false;
               UI_Display_Update_Options(false);
            }
            else
            {
               DisplayAdminText("Please enter the first and last name of the user to update");
            }
         }
         catch (Exception)
         {

            throw;
         }
      }

      private void Setup_Midnight_Timer()
      {
         TimeSpan timeToGo = new TimeSpan(1, 0, 0, 0) - DateTime.Now.TimeOfDay; //new DateTime(2018, 12, 25, 0, 0, 0);
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
         try
         {
            gAPI = new GSheetsAPI();
            gAPI.AuthorizeGoogleApp();
            Log("Application authorized");
            Log("Refreshing local data");
            if (gAPI.Refresh_Local_Data())
            {
               Log("Successfully refreshed local data");
               UI_Control(true);
               Setup_Midnight_Timer();
               TXT_ID_Scan.Focus(FocusState.Keyboard);
            }
            else
            {
               Log("Failed to refresh local data");
               UI_Control(false);
               BTN_Reconnect.Focus(FocusState.Keyboard);
            }
         }
         catch (Exception)
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
            TXT_ID_Scan.IsEnabled = true;
            TXT_ID_Scan.Visibility = Visibility.Visible;
            Log("Disabling reconnect button");
            GRD_Admin.Visibility = Visibility.Collapsed;
            BTN_Save_Updated_User.IsEnabled = false;
            BTN_Save_Updated_User.Visibility = Visibility.Collapsed;
            BTN_Reconnect.Visibility = Visibility.Collapsed;
            BTN_Reconnect.IsEnabled = false;

         }
         else
         {
            Log("Disabling text entry field");
            TXT_ID_Scan.IsEnabled = false;
            TXT_ID_Scan.Visibility = Visibility.Collapsed;
            Log("Enabling reconnect button");
            GRD_Admin.Visibility = Visibility.Collapsed;
            BTN_Reconnect.Visibility = Visibility.Visible;
            BTN_Reconnect.IsEnabled = true;
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
            BTN_Save_Updated_User.Visibility = Visibility.Collapsed;
         }
      }

      private bool Lookup_ID(ulong ID)
      {
         bool
             success = false;

         try
         {
            Log(string.Format("Verifying ID: {0}", ID.ToString()));
            if (gAPI.Check_Valid_ID(ID))
            {
               Log(string.Format("ID: {0} - Verified", ID.ToString()));
               success = true;
            }
            else
            {
               Log(string.Format("ID: {0} - Not registered", ID.ToString()));
            }
         }
         catch (Exception ex)
         {
            HandleException(ex, MethodBase.GetMethodFromHandle(new RuntimeMethodHandle()).Name);
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
               success = true;
            }
            else
            {
               Log(string.Format("ID: {0} - Unauthorized", ID.ToString()));
            }
         }
         catch (Exception ex)
         {
            HandleException(ex, MethodBase.GetMethodFromHandle(new RuntimeMethodHandle()).Name);
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
            HandleException(ex, MethodBase.GetMethodFromHandle(new RuntimeMethodHandle()).Name);
         }
      }

      private void Check_Out_Users()
      {
         try
         {
            //get list of users still checked in
            List<User> stillCheckedInUsers = gAPI.Get_CheckedIn_Users();
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
         }
         catch (Exception ex)
         {
            HandleException(ex, MethodBase.GetMethodFromHandle(new RuntimeMethodHandle()).Name);
         }
         finally
         {
            Setup_Midnight_Timer();
         }
      }

      internal void Log(string text)
      {
         _log.Info(text);
      }

      #endregion

      #region *** DISPLAY/EXCEPTION HANDLING ***
      internal async void DisplayText(string text)
      {
         if (!txt_Status.Dispatcher.HasThreadAccess)
         {
            await Dispatcher.RunAsync(Windows.UI.Core.CoreDispatcherPriority.Normal, () => 
            {
               txt_Status.Text = text;
               System.Threading.Tasks.Task.Delay(1500);
               txt_Status.Text = string.Empty;
            });
         }
         else
         {
            txt_Status.Text = text;
            await System.Threading.Tasks.Task.Delay(1500);
            txt_Status.Text = string.Empty;
         }
      }

      internal void DisplayAdminText(string text)
      {
         Paragraph paragraph = new Paragraph();
         Run run = new Run();

         // Customize some properties on the RichTextBlock.
         RTB_AdminOutput.IsTextSelectionEnabled = true;
         RTB_AdminOutput.SelectionHighlightColor = new SolidColorBrush(Windows.UI.Colors.Pink);
         RTB_AdminOutput.Foreground = new SolidColorBrush(Windows.UI.Colors.Blue);
         RTB_AdminOutput.FontWeight = Windows.UI.Text.FontWeights.Light;
         RTB_AdminOutput.FontFamily = new FontFamily("Arial");
         RTB_AdminOutput.FontStyle = Windows.UI.Text.FontStyle.Normal;
         run.Text = text;

         //Add the Run to the Paragraph, the Paragraph to the RichTextBlock.
         paragraph.Inlines.Add(run);
         RTB_AdminOutput.Blocks.Add(paragraph);
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

      private void RTB_AdminOutput_SelectionChanged(object sender, RoutedEventArgs e)
      {

      }
   }
}
