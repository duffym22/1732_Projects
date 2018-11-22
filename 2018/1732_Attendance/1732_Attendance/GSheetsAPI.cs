using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
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
        const string _OUT_STATUS = "OUT";
        const string _IN_STATUS = "IN";
        const string _ATTENDANCE_STATUS_RANGE = "ATTENDANCE_STATUS!A:C";
        const string _ATTENDANCE_STAT_ID_RANGE = "ATTENDANCE_STATUS!A:A";
        const string _ATTENDANCE_STAT_NAME_RANGE = "ATTENDANCE_STATUS!B:B";
        const string _LOG_RANGE = "LOG!A:C";


        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        SheetsService service;
        UserCredential credential;
        StringBuilder _exMsg;

        string[] Scopes = { SheetsService.Scope.Spreadsheets };
        string ApplicationName = "1732 Attendance Check-In Station";
        string SheetId = "13U-gYgtXlh8Q0Qgaim6nzrFlkOAP4dJP2hvOTaO7nTg";
        Dictionary<int, List<string>> dict_Attendance;

        #endregion

        #region *** PROPERTIES ***
        public string LastException { get { return _exMsg.ToString(); } }
        #endregion

        #region *** CONSTRUCTOR ***

        public GSheetsAPI()
        {
            service = new SheetsService();
            credential = null;
            dict_Attendance = new Dictionary<int, List<string>>();
        }

        #endregion

        #region *** FEATURE FUNCTIONALITY METHODS ***

        public void Add_User(int ID, string name)
        {
            try
            {
                InsertRows(Create_Attendance_Status_Row(ID, name, _OUT_STATUS), SheetId, _ATTENDANCE_STATUS_RANGE, service);
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
                Read_Attendance_Status();
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
                string rowToUpdate = Get_ID_Status_Row(ID);
                UpdateRows(Create_Attendance_Status_Row(ID, dict_Attendance[ID][0],
                    dict_Attendance[ID][1].Equals(_IN_STATUS)
                    ? _OUT_STATUS
                    : _IN_STATUS)
                    , SheetId, rowToUpdate, service);

                Read_Attendance_Status();
                InsertRows(Create_Log_Row(ID, dict_Attendance[ID][1]), SheetId, Get_Next_Log_Row(), service);
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        #endregion

        #region *** GET/READ METHODS ***
        private void Read_Attendance_Status()
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

        private string Get_Next_Log_Row()
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
                if (dict_Attendance.ContainsKey(checkID))
                {
                    Console.WriteLine("Exists!");
                    Console.WriteLine("Name: " + dict_Attendance[checkID][0]);
                    Console.WriteLine("Status: " + dict_Attendance[checkID][1]);
                    do
                    {
                        Console.WriteLine("Would you like to change the status?");
                        read = Console.ReadLine().Trim().ToUpper();
                    }
                    while (!read.Equals("Y") && !read.Equals("N"));

                    if (read.Equals("Y"))
                    {
                        string rowToUpdate = Get_ID_Status_Row(checkID);
                        UpdateRows(
                            Create_Attendance_Status_Row(checkID, dict_Attendance[checkID][0],
                            dict_Attendance[checkID][1].Equals(_IN_STATUS)
                            ? _OUT_STATUS
                            : _IN_STATUS)
                            , SheetId, rowToUpdate, service);
                        Read_Attendance_Status();
                        InsertRows(Create_Log_Row(checkID, dict_Attendance[checkID][1]), SheetId, Get_Next_Log_Row(), service);
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

        private void Parse_Attendance_Status_Rows(IList<IList<object>> idStatusList)
        {
            string 
                name, 
                stat;

            try
            {
                //Start at 1 because first row is the header (ID | Name)
                for (int i = 1; i < idStatusList.Count; i++)
                {
                    //Get the current row (ID | Current_Status)
                    IList<object> row = idStatusList[i];
                    int.TryParse((string)row[0], out int numId);
                    name = row[1].ToString();
                    stat = row[2].ToString();

                    //If the dictionary has already been populated, then just update the value
                    if (dict_Attendance.ContainsKey(numId))
                    {
                        dict_Attendance[numId][1] = stat;
                    }
                    else
                    {
                        dict_Attendance.Add(numId, new List<string> { name, stat });
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

        private IList<IList<object>> Create_Log_Row(int ID, string status)
        {
            IList<IList<object>> newRow = new List<IList<object>>();
            newRow.Add(new List<object>() { ID, DateTime.Now.ToString(), status.ToString() });
            return newRow;
        }

        private IList<IList<object>> Create_Attendance_Status_Row(int ID, string name, string status)
        {
            IList<IList<object>> newRow = new List<IList<object>>();
            newRow.Add(new List<object>() { ID, name, status.ToString() });
            return newRow;
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

        private void DeleteRows(int rowToDelete, string spreadsheetId, string newRange, SheetsService service)
        {
            //DELETE THIS ROW
            Request RequestBody = new Request()
            {
                DeleteDimension = new DeleteDimensionRequest()
                {
                    Range = new DimensionRange()
                    {
                        SheetId = 0,
                        Dimension = "ROWS",
                        StartIndex = rowToDelete,
                        EndIndex = rowToDelete + 1
                    }
                }
            };

            List<Request> RequestContainer = new List<Request> { RequestBody };

            BatchUpdateSpreadsheetRequest DeleteRequest = new BatchUpdateSpreadsheetRequest();
            DeleteRequest.Requests = RequestContainer;

            SpreadsheetsResource.BatchUpdateRequest Deletion = new SpreadsheetsResource.BatchUpdateRequest(service, DeleteRequest, spreadsheetId);
            Deletion.Execute();
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
        }
        #endregion
    }
}
