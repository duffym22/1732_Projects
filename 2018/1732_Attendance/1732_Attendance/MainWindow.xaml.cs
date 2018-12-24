using log4net;
using log4net.Config;
using System;
using System.Collections.Generic;
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
      internal ulong ID_Scan;
      private GSheetsAPI gAPI;
      private static readonly ILog _log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

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

      private void BTN_Login_Click(object sender, RoutedEventArgs e)
      {

      }

      private void BTN_Reconnect_Click(object sender, RoutedEventArgs e)
      {
         Initialize();
      }

      private void BTN_Add_User_Click(object sender, RoutedEventArgs e)
      {
         try
         {
            if (!string.IsNullOrEmpty(TXT_ID.Text) && !string.IsNullOrEmpty(TXT_Name.Text))
            {

               if (TXT_ID.Text.Length > 10)
               {
                  string shortID = TXT_ID.Text.Substring(txt_ID_Scan.Text.Length / 2, (TXT_ID.Text.Length - (TXT_ID.Text.Length / 2) - 1));
                  TXT_ID.Text = shortID;
               }

               ulong.TryParse(TXT_ID.Text, out ulong ID);
               if (gAPI.Add_User(ID, TXT_Name.Text, (bool)CHK_Is_Mentor.IsChecked))
               {
                  DisplayText(string.Format("Successfully added ID: {0} | NAME: {1}", TXT_ID.Text, TXT_Name.Text));
                  Log(string.Format("Added ID: {0} | NAME: {1}", TXT_ID.Text, TXT_Name.Text));
               }
               else
               {
                  DisplayText(string.Format("Failed to add ID: {0} | NAME: {1}", TXT_ID.Text, TXT_Name.Text));
                  Log(string.Format("Failed to add ID: {0} | NAME: {1}", TXT_ID.Text, TXT_Name.Text));
                  Log(gAPI.LastException);
               }
               TXT_ID.Clear();
               TXT_Name.Clear();
               CHK_Is_Mentor.IsChecked = false;
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
            if (!string.IsNullOrEmpty(TXT_ID.Text))
            {
               if (TXT_ID.Text.Length > 10)
               {
                  string shortID = TXT_ID.Text.Substring(txt_ID_Scan.Text.Length / 2, (TXT_ID.Text.Length - (TXT_ID.Text.Length / 2) - 1));
                  TXT_ID.Text = shortID;
               }

               ulong.TryParse(TXT_ID.Text, out ulong ID);
               if (gAPI.Delete_User(ID))
               {
                  DisplayText(string.Format("Successfully deleted ID: {0}", TXT_ID.Text));
                  Log(string.Format("Deleted ID: {0}", TXT_ID.Text));
               }
               else
               {
                  DisplayText(string.Format("Failed to delete ID: {0}", TXT_ID.Text));
                  Log(string.Format("Did not find ID: {0}", TXT_ID.Text));
                  Log(gAPI.LastException);
               }
               TXT_ID.Clear();
            }
         }
         catch (Exception ex)
         {
            HandleException(ex, MethodBase.GetCurrentMethod().Name);
         }
      }

      private void BTN_Who_CheckedIn_Click(object sender, RoutedEventArgs e)
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
               Log(string.Format("Displayed currently logged in users. Count = {0}", users.Count));
               UserDataGrid.ItemsSource = users;
            }
            else
            {
               Log("No users currently logged in");
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
               UserDataGrid.ItemsSource = users;
            }
            else
            {
               Log("No users recently logged out");
            }
         }
         catch (Exception ex)
         {
            HandleException(ex, MethodBase.GetCurrentMethod().Name);
         }
      }

      private void BTN_Refresh_Click(object sender, RoutedEventArgs e)
      {
         try
         {
            if (gAPI.Refresh_Local_Data())
            {
               Log("Local data refreshed");
            }
            else
            {
               Log("Failed to refresh local data");
               Log(gAPI.LastException);
            }
         }
         catch (Exception ex)
         {
            HandleException(ex, MethodBase.GetCurrentMethod().Name);
         }
      }

      private void TXT_ID_Scan_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
      {
         try
         {
            if (e.Key == System.Windows.Input.Key.Enter || e.Key == System.Windows.Input.Key.Return)
            {
               if (txt_ID_Scan.Text.Length > 10)
               {
                  string shortID = txt_ID_Scan.Text.Substring(txt_ID_Scan.Text.Length / 2, (txt_ID_Scan.Text.Length - (txt_ID_Scan.Text.Length / 2) - 1));
                  txt_ID_Scan.Text = shortID;
               }

               if (ulong.TryParse(txt_ID_Scan.Text, out ID_Scan))
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

      private void UserDataGrid_LostFocus(object sender, RoutedEventArgs e)
      {
         UserDataGrid.Visibility = Visibility.Hidden;
      }

      private void UserDataGrid_FocusableChanged(object sender, DependencyPropertyChangedEventArgs e)
      {
         ///idk something
         ///
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
               BTN_Reconnect.Focus();
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
            BTN_Reconnect.Visibility = Visibility.Hidden;
            BTN_Reconnect.IsEnabled = false;

         }
         else
         {
            Log("Disabling text entry field");
            txt_ID_Scan.IsEnabled = false;
            txt_ID_Scan.Visibility = Visibility.Hidden;
            Log("Enabling reconnect button");
            BTN_Reconnect.Visibility = Visibility.Visible;
            BTN_Reconnect.IsEnabled = true;
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
