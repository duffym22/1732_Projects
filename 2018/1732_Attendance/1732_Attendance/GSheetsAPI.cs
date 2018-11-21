using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource;

namespace _1732_Attendance
{
    /// ***** PSC: Normal Use *****
    /// 1. Scan ID over sensor
    /// 2. Text populates on field - automatic carriage return will invoke keypress
    /// 3. Check for carriage return in keypress event and if present, continue check
    /// 4. Parse ID from field as integer and check against existing dictionary read on startup of app (or from periodic invoke)
    /// 5. If ID is in dictionary, create new entry to record to LOG tab the ID and timestamp. 
    /// 5a. If ID is NOT in the directory, display to screen "ID: [IDVAL] is not in the list of valid IDs. Please contact a mentor to be added"
    /// 6. Read the current status' of all IDs from the ATTENDANCE_STATUS tab
    /// 7. Enumerate current status of all IDs into dict_ID_Status
    /// 8. Verify current status of the ID and invert it to write to the ATTENDANCE_STATUS tab
    class GSheetsAPI
    {
        #region *** FIELDS ***
        const string _IDS_RANGE = "IDS!A:B";
        const string _ATTENDANCE_STATUS_RANGE = "ATTENDANCE_STATUS!A:B";
        const string _ATTENDANCE_STAT_ID_RANGE = "ATTENDANCE_STATUS!A:A";
        const string _LOG_RANGE = "LOG!A:C";


        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        SheetsService service;
        UserCredential credential;
        StringBuilder _exMsg;

        string[] Scopes = { SheetsService.Scope.Spreadsheets };
        string ApplicationName = "1732 Attendance Check-In Station";
        string SheetId = "13U-gYgtXlh8Q0Qgaim6nzrFlkOAP4dJP2hvOTaO7nTg";
        Dictionary<int, string> dict_ID_Name;
        Dictionary<int, ATTENDANCE_STATUS> dict_ID_Status;

        #endregion

        #region *** PROPERTIES ***
        public string LastException { get { return _exMsg.ToString(); } }
        #endregion

        #region *** ENUMS/STRUCTURES ***
        enum ATTENDANCE_STATUS
        {
            IN,
            OUT,
            UNKNOWN
        }

        #endregion

        #region *** CONSTRUCTOR ***

        public GSheetsAPI()
        {
            service = new SheetsService();
            credential = null;
            dict_ID_Name = new Dictionary<int, string>();
            dict_ID_Status = new Dictionary<int, ATTENDANCE_STATUS>();
        }

        #endregion
        
        #region *** FEATURE FUNCTIONALITY METHODS ***

        public void Add_User(int ID)
        {
            try
            {
                InsertRows(Create_ID_Status_Row(ID, ATTENDANCE_STATUS.OUT), SheetId, _ATTENDANCE_STATUS_RANGE, service);
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        public void Update_User()
        {
            try
            {

            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        public void Delete_User()
        {
            try
            {

            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        public void Refresh_Local_Data()
        {
            try
            {
                Get_Current_ID_List();
                Get_Current_ID_Status();
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        public void Update_User_Status()
        {
            try
            {

            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        #endregion

        #region *** GET METHODS ***
        private void Get_Current_ID_List()
        {
            try
            {
                GetRequest getRequest = service.Spreadsheets.Values.Get(SheetId, _IDS_RANGE);

                ValueRange getResponse = getRequest.Execute();
                IList<IList<Object>> idList = getResponse.Values;

                Parse_IDS_Name(idList);
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        private void Get_Current_ID_Status()
        {
            try
            {
                GetRequest getRequest = service.Spreadsheets.Values.Get(SheetId, _ATTENDANCE_STATUS_RANGE);

                ValueRange getResponse = getRequest.Execute();
                IList<IList<Object>> idList = getResponse.Values;

                Parse_IDS_Status(idList);
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        private string Get_ID_Status_Row(int ID)
        {
            string returnVal = string.Empty;
            int readID;
            try
            {
                GetRequest getRequest = service.Spreadsheets.Values.Get(SheetId, _ATTENDANCE_STAT_ID_RANGE);

                ValueRange getResponse = getRequest.Execute();
                IList<IList<Object>> getValues = getResponse.Values;

                if (getValues != null)
                {
                    for (int i = 0; i < getValues.Count; i++)
                    {
                        IList<object> row = getValues[i];
                        int.TryParse(row[0].ToString(), out readID);
                        if (readID.Equals(ID))
                        {
                            Console.WriteLine("Found it! It's on row: " + i);
                            returnVal = "ATTENDANCE_STATUS!A" + (i + 1) + ":B" + (i + 1);         //row to be updated
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return returnVal;
        }

        private string Get_NextAttendanceRow()
        {
            string returnVal = string.Empty;
            try
            {
                GetRequest getRequest = service.Spreadsheets.Values.Get(SheetId, _LOG_RANGE);

                ValueRange getResponse = getRequest.Execute();
                IList<IList<Object>> getValues = getResponse.Values;

                returnVal = string.Format("LOG!A{0}:C", getValues.Count + 1);
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return returnVal;
        }
        #endregion

        #region *** CHECK/PARSE METHODS ***
        private void CheckID(int checkID)
        {
            string read;
            try
            {
                if (dict_ID_Name.ContainsKey(checkID) && dict_ID_Status.ContainsKey(checkID))
                {
                    Console.WriteLine("Exists!");
                    Console.WriteLine("Name: " + dict_ID_Name[checkID]);
                    Console.WriteLine("Status: " + dict_ID_Status[checkID]);
                    do
                    {
                        Console.WriteLine("Would you like to change the status?");
                        read = Console.ReadLine().Trim().ToUpper();
                    }
                    while (!read.Equals("Y") && !read.Equals("N"));

                    if (read.Equals("Y"))
                    {
                        string rowToUpdate = Get_ID_Status_Row(checkID);
                        UpdateRows(Create_ID_Status_Row(checkID,
                            dict_ID_Status[checkID].Equals(ATTENDANCE_STATUS.IN)
                            ? ATTENDANCE_STATUS.OUT
                            : ATTENDANCE_STATUS.IN)
                            , SheetId, rowToUpdate, service);
                        Get_Current_ID_Status();
                        InsertRows(Create_ID_Timestamp_Row(checkID, dict_ID_Status[checkID]), SheetId, Get_NextAttendanceRow(), service);
                    }
                }
                else if (dict_ID_Name.ContainsKey(checkID))
                {
                    Console.WriteLine("Only exists in ID|Name list!");
                    Console.WriteLine("Name: " + dict_ID_Name[checkID]);

                    do
                    {
                        Console.WriteLine("Add to ID|STATUS List?");
                        read = Console.ReadLine().Trim().ToUpper();
                    }
                    while (!read.Equals("Y") && !read.Equals("N"));

                    if (read.Equals("Y"))
                    {
                        InsertRows(Create_ID_Status_Row(checkID, ATTENDANCE_STATUS.OUT), SheetId, _ATTENDANCE_STATUS_RANGE, service);
                    }
                }
                else if (dict_ID_Status.ContainsKey(checkID))
                {
                    Console.WriteLine("Only exists in ID|STATUS list!");
                    Console.WriteLine("Status: " + dict_ID_Status[checkID]);
                    do
                    {
                        Console.WriteLine("Add to ID|Name List?");
                        read = Console.ReadLine().Trim().ToUpper();
                    }
                    while (!read.Equals("Y") && !read.Equals("N"));

                    if (read.Equals("Y"))
                    {
                        Console.WriteLine("Please enter a name:");
                        read = Console.ReadLine();
                        InsertRows(Create_ID_Name_Row(checkID, read), SheetId, _IDS_RANGE, service);

                    }
                }
                else
                {
                    Console.WriteLine("Didn't find that ID in any list");
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        private void Parse_IDS_Name(IList<IList<object>> idNameList)
        {
            string name = string.Empty;
            int numId = 0;
            try
            {
                //Start at 1 because first row is the header (ID | Name)
                for (int i = 1; i < idNameList.Count; i++)
                {
                    name = string.Empty;
                    numId = 0;
                    IList<object> row = idNameList[i];
                    if (row.Count.Equals(2))
                    {

                        int.TryParse((string)row[0], out numId);
                        name = (string)row[1];
                        dict_ID_Name.Add(numId, name);
                    }
                    else
                    {
                        Console.WriteLine("Only ID found for row " + i + 1);
                    }
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        private void Parse_IDS_Status(IList<IList<object>> idStatusList)
        {
            ATTENDANCE_STATUS stat;
            try
            {
                //Start at 1 because first row is the header (ID | Name)
                for (int i = 1; i < idStatusList.Count; i++)
                {
                    stat = ATTENDANCE_STATUS.UNKNOWN;
                    //Get the current row (ID | Current_Status)
                    IList<object> row = idStatusList[i];
                    int numId;
                    int.TryParse((string)row[0], out numId);
                    try { stat = (ATTENDANCE_STATUS)Enum.Parse(typeof(ATTENDANCE_STATUS), row[1].ToString(), true); }
                    catch { }

                    //If the dictionary has already been populated, then just update the value
                    if (dict_ID_Status.ContainsKey(numId))
                    {
                        dict_ID_Status[numId] = stat;
                    }
                    else
                    {
                        dict_ID_Status.Add(numId, stat);
                    }
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        #endregion

        #region *** RECORD CREATION METHODS ***

        private IList<IList<object>> Create_ID_Name_Row(int ID, string name)
        {
            IList<IList<object>> newRow = new List<IList<object>>();
            newRow.Add(new List<object>() { ID, name });
            return newRow;
        }

        private IList<IList<object>> Create_ID_Timestamp_Row(int ID, ATTENDANCE_STATUS status)
        {
            IList<IList<object>> newRow = new List<IList<object>>();
            newRow.Add(new List<object>() { ID, DateTime.Now.ToString(), status.ToString() });
            return newRow;
        }

        private IList<IList<object>> Create_ID_Status_Row(int ID, ATTENDANCE_STATUS status)
        {
            IList<IList<object>> newRow = new List<IList<object>>();
            newRow.Add(new List<object>() { ID, status.ToString() });
            return newRow;
        }

        private IList<IList<object>> CreateRecord()
        {
            IList<IList<object>> appRows = new List<IList<object>>();
            IList<object> vals = new List<object>() { "2222", DateTime.UtcNow.ToString(), "IN" };
            appRows.Add(vals);
            return appRows;
        }

        #endregion

        #region *** CONNECT/WRITE/UPDATE DATA METHODS ***
        public bool AuthorizeGoogleApp()
        {
            bool success = false;
            try
            {
                using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                {
                    string credPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                    credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");

                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                    Console.WriteLine("Credential file saved to: " + credPath);
                }

                // Create Google Sheets API service.
                service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });
                success = true;
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return success;
        }

        private void InsertRows(IList<IList<Object>> values, string spreadsheetId, string newRange, SheetsService service)
        {
            AppendRequest request = service.Spreadsheets.Values.Append(new ValueRange() { Values = values }, spreadsheetId, newRange);
            request.InsertDataOption = AppendRequest.InsertDataOptionEnum.INSERTROWS;
            request.ValueInputOption = AppendRequest.ValueInputOptionEnum.RAW;
            var response = request.Execute();
        }

        private void UpdateRows(IList<IList<Object>> values, string spreadsheetId, string newRange, SheetsService service)
        {
            UpdateRequest request = service.Spreadsheets.Values.Update(new ValueRange() { Values = values }, spreadsheetId, newRange);
            request.ValueInputOption = UpdateRequest.ValueInputOptionEnum.RAW;
            var response = request.Execute();
        }

        #endregion

        #region *** EXCEPTION/GUI HANDLING ***
        private void HandleException(Exception ex, string callingMethod)
        {
            _exMsg = new StringBuilder();

            _exMsg.AppendLine(string.Format("Exception thrown in: {0}", callingMethod));
            _exMsg.AppendLine(string.IsNullOrEmpty(ex.Message) ? "" : ex.Message);
            _exMsg.AppendLine(string.IsNullOrEmpty(ex.Source) ? "" : ex.Source);
            _exMsg.AppendLine(string.IsNullOrEmpty(ex.StackTrace.ToString()) ? "" : ex.StackTrace.ToString());

            Console.WriteLine("");
            Console.WriteLine(_exMsg);
            Console.WriteLine("");
        }
        #endregion
    }
}
