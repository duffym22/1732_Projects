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

namespace Test_GoogleSheetsAPI
{
    class Sheets_Append
    {
        const string _ADDED_STATUS = "ADDED";
        const string _DELETED_STATUS = "DELETED";
        const string _OUT_STATUS = "OUT";
        const string _IN_STATUS = "IN";
        const string _ATTENDANCE_STATUS_RANGE = "ATTENDANCE_STATUS!A:D";
        const string _ATTENDANCE_STAT_ID_RANGE = "ATTENDANCE_STATUS!A:A";
        const string _ATTENDANCE_STAT_NAME_RANGE = "ATTENDANCE_STATUS!B:B";
        const string _LOG_RANGE = "LOG!A:C";

        const int _ATTENDANCE_STATUS_GID = 741019777;
        const int _ATTENDANCE_LOG_GID = 1617370344;

        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static SheetsService service;
        static UserCredential credential;
        static StringBuilder _exMsg;
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "1732 Attendance Check-In Station";
        static readonly string SheetId = "13U-gYgtXlh8Q0Qgaim6nzrFlkOAP4dJP2hvOTaO7nTg";
        static Dictionary<int, List<string>> dict_Attendance = new Dictionary<int, List<string>>();

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
        /// 9. 

        static void Main(string[] args)
        {
            Stopwatch watch = new Stopwatch();
            string 
                cmd = string.Empty, 
                name = string.Empty;

            int ID = -1;

            watch.Start();
            AuthorizeGoogleApp();
            watch.Stop();
            Console.WriteLine(string.Format("T2-AUTH: {0}ms", watch.ElapsedMilliseconds));

            watch.Restart();
            Read_Attendance_Status();
            watch.Stop();
            Console.WriteLine(string.Format("T2-READ: {0}ms", watch.ElapsedMilliseconds));

            do
            {
                Console.WriteLine("Enter a command");
                Console.WriteLine("A = Add User | C = Change User Status | D = Delete User | R = Read User Status | Q = Quit");
                cmd = Console.ReadLine().Trim().ToUpper();
                if (!cmd.Equals("Q"))
                {
                    Console.WriteLine("Enter an ID of a user");
                    int.TryParse(Console.ReadLine(), out ID);
                    if (!dict_Attendance.ContainsKey(ID))
                    {
                        Console.WriteLine("ID does not exist");
                        cmd = "X";
                    }

                    if (cmd.Equals("A"))
                    {
                        Console.WriteLine("Please enter a name in the following format: Last Name, First Name");
                        name = Console.ReadLine();
                    }
                }

                watch.Restart();
                switch (cmd)
                {
                    case "A":
                        InsertRows(Create_Attendance_Status_Row(ID, name, "OUT"), SheetId, string.Format("ATTENDANCE_STATUS!A{0}:D", Get_Next_Attendance_Row()), service);
                        Read_Attendance_Status();
                        InsertRows(Create_Log_Row(ID, _ADDED_STATUS), SheetId, Get_Next_Log_Row(), service);
                        break;
                    case "C":
                        int rowToUpdate = Get_Attendance_Status_Row(ID);
                        string rowRange = "ATTENDANCE_STATUS!A" + (rowToUpdate + 1) + ":D" + (rowToUpdate + 1);         //row to be updated - increment by 1 because sheets start at "0"
                        UpdateRows(
                            Create_Attendance_Status_Row(ID, dict_Attendance[ID][0],
                            dict_Attendance[ID][1].Equals(_IN_STATUS)
                            ? _OUT_STATUS
                            : _IN_STATUS)
                            , SheetId, rowRange, service);
                        Read_Attendance_Status();
                        InsertRows(Create_Log_Row(ID, dict_Attendance[ID][1]), SheetId, Get_Next_Log_Row(), service);
                        break;
                    case "D":
                        int rowToRemove = Get_Attendance_Status_Row(ID);
                        DeleteRows(rowToRemove, SheetId, service);
                        Read_Attendance_Status();
                        InsertRows(Create_Log_Row(ID, _DELETED_STATUS), SheetId, Get_Next_Log_Row(), service);
                        break;
                    case "R":
                        if (dict_Attendance.ContainsKey(ID))
                        {
                            Console.WriteLine("Name: " + dict_Attendance[ID][0]);
                            Console.WriteLine("Status: " + dict_Attendance[ID][1]);
                            Console.WriteLine("Last Updated: " + dict_Attendance[ID][2]);
                        }
                        else
                        {
                            Console.WriteLine("ID does not exist");
                        }

                        break;
                }
                watch.Stop();
                Console.WriteLine(string.Format("T2-CMD: {0}ms", watch.ElapsedMilliseconds));
            }
            while (!cmd.Equals("Q"));
        }

        #region *** FEATURE FUNCTIONALITY METHODS ***

        public void Add_User(int ID, string name)
        {
            try
            {
                InsertRows(Create_Attendance_Status_Row(ID, name, "OUT"), SheetId, string.Format("ATTENDANCE_STATUS!A{0}:D", Get_Next_Attendance_Row()), service);
                Read_Attendance_Status();
                InsertRows(Create_Log_Row(ID, _ADDED_STATUS), SheetId, Get_Next_Log_Row(), service);
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        public void Update_User_Status(int ID)
        {
            try
            {
                int rowToUpdate = Get_Attendance_Status_Row(ID);
                string rowRange = "ATTENDANCE_STATUS!A" + (rowToUpdate + 1) + ":D" + (rowToUpdate + 1);         //row to be updated - increment by 1 because sheets start at "0"
                UpdateRows(
                    Create_Attendance_Status_Row(ID, dict_Attendance[ID][0],
                    dict_Attendance[ID][1].Equals(_IN_STATUS)
                    ? _OUT_STATUS
                    : _IN_STATUS)
                    , SheetId,
                    rowRange,
                    service);
                Read_Attendance_Status();
                InsertRows(Create_Log_Row(ID, dict_Attendance[ID][1]), SheetId, Get_Next_Log_Row(), service);
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        public void Delete_User(int ID)
        {
            try
            {
                int rowToRemove = Get_Attendance_Status_Row(ID);
                DeleteRows(rowToRemove, SheetId, service);
                Read_Attendance_Status();
                InsertRows(Create_Log_Row(ID, _DELETED_STATUS), SheetId, Get_Next_Log_Row(), service);
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
                Read_Attendance_Status();
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        #endregion

        #region *** GET/READ METHODS ***
        private static void Read_Attendance_Status()
        {
            try
            {
                GetRequest getRequest = service.Spreadsheets.Values.Get(SheetId, _ATTENDANCE_STATUS_RANGE);

                ValueRange getResponse = getRequest.Execute();
                IList<IList<Object>> idList = getResponse.Values;

                Parse_Attendance_Status_Rows(idList);
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        private static int Get_Attendance_Status_Row(int ID)
        {
            int returnVal = -1;
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
                        int.TryParse(row[0].ToString(), out int readID);
                        if (readID.Equals(ID))
                        {
                            returnVal = i;
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

        private static int Get_Next_Attendance_Row()
        {
            int returnVal = -1;
            try
            {
                GetRequest getRequest = service.Spreadsheets.Values.Get(SheetId, _ATTENDANCE_STATUS_RANGE);

                ValueRange getResponse = getRequest.Execute();
                IList<IList<Object>> getValues = getResponse.Values;

                returnVal = getValues.Count + 1;
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return returnVal;
        }

        private static string Get_Next_Log_Row()
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

        #region *** PARSE METHODS ***
        private static void Parse_Attendance_Status_Rows(IList<IList<object>> idStatusList)
        {
            string
                name,
                stat;

            try
            {
                //Wipe any previous dictionary values to start fresh with every request
                //Treats the Google Sheet as the golden copy
                dict_Attendance = new Dictionary<int, List<string>>();
                //Start at 1 because first row is the header (ID | Name)
                for (int i = 1; i < idStatusList.Count; i++)
                {
                    //Get the current row (ID | Current_Status)
                    IList<object> row = idStatusList[i];
                    int.TryParse((string)row[0], out int ID);
                    name = row[1].ToString();
                    stat = row[2].ToString();
                    DateTime.TryParse(row[3].ToString(), out DateTime lastUpdated);
                    dict_Attendance.Add(ID, new List<string> { name, stat, lastUpdated.ToString() });
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        #endregion

        #region *** RECORD CREATION METHODS ***

        private static IList<IList<object>> Create_Log_Row(int ID, string status)
        {
            string log = string.Empty;
            switch (status)
            {
                case _IN_STATUS:
                    log = "User checked IN";
                    break;
                case _OUT_STATUS:
                    log = "User checked OUT";
                    break;
                case _ADDED_STATUS:
                    log = "User ADDED";
                    break;
                case _DELETED_STATUS:
                    log = "User DELETED";
                    break;
                default:
                    throw new Exception("Unable to parse status to create row for Log sheet");
            }
            IList<IList<object>> newRow = new List<IList<object>>
            {
                new List<object>() { ID, DateTime.Now.ToString(), log }
            };
            return newRow;
        }

        private static IList<IList<object>> Create_Attendance_Status_Row(int ID, string name, string status)
        {
            IList<IList<object>> newRow = new List<IList<object>>
            {
                new List<object>() { ID, name, status, DateTime.Now.ToString() }
            };
            return newRow;
        }

        #endregion

        #region *** CONNECT/WRITE/UPDATE DATA METHODS ***
        public static bool AuthorizeGoogleApp()
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

        private static void InsertRows(IList<IList<Object>> values, string spreadsheetId, string newRange, SheetsService service)
        {
            AppendRequest request = service.Spreadsheets.Values.Append(new ValueRange() { Values = values }, spreadsheetId, newRange);
            request.InsertDataOption = AppendRequest.InsertDataOptionEnum.INSERTROWS;
            request.ValueInputOption = AppendRequest.ValueInputOptionEnum.RAW;
            AppendValuesResponse response = request.Execute();
        }

        private static void UpdateRows(IList<IList<Object>> values, string spreadsheetId, string newRange, SheetsService service)
        {
            UpdateRequest request = service.Spreadsheets.Values.Update(new ValueRange() { Values = values }, spreadsheetId, newRange);
            request.ValueInputOption = UpdateRequest.ValueInputOptionEnum.RAW;
            UpdateValuesResponse response = request.Execute();
        }

        private static void DeleteRows(int rowToDelete, string spreadsheetId, SheetsService service)
        {
            BatchUpdateSpreadsheetRequest content = new BatchUpdateSpreadsheetRequest();
            Request request = new Request()
            {
                DeleteDimension = new DeleteDimensionRequest()
                {
                    Range = new DimensionRange()
                    {
                        SheetId = _ATTENDANCE_STATUS_GID,
                        Dimension = "ROWS",
                        StartIndex = rowToDelete,
                        EndIndex = rowToDelete + 1
                    }
                }
            };

            List<Request> listRequests = new List<Request> { request };
            content.Requests = listRequests;

            SpreadsheetsResource.BatchUpdateRequest Deletion = new SpreadsheetsResource.BatchUpdateRequest(service, content, spreadsheetId);
            BatchUpdateSpreadsheetResponse response = Deletion.Execute();
        }

        #endregion

        #region *** EXCEPTION/GUI HANDLING ***
        private static void HandleException(Exception ex, string callingMethod)
        {
            _exMsg = new StringBuilder();

            _exMsg.AppendLine(string.Format("Exception thrown in: {0}", callingMethod));
            _exMsg.AppendLine(string.IsNullOrEmpty(ex.Message) ? "" : ex.Message);
            _exMsg.AppendLine(string.IsNullOrEmpty(ex.Source) ? "" : ex.Source);
            _exMsg.AppendLine(string.IsNullOrEmpty(ex.StackTrace.ToString()) ? "" : ex.StackTrace.ToString());

            Console.Write(_exMsg.ToString());
        }
        #endregion
    }
}
