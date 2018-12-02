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
        const string _ADDED_STATUS = "ADDED";
        const string _DELETED_STATUS = "DELETED";
        const string _OUT_STATUS = "OUT";
        const string _IN_STATUS = "IN";
        const string _ATTENDANCE_STATUS_RANGE = _ATTENDANCE_STATUS_START_RANGE + ":" + _ATTENDANCE_STATUS_END_RANGE;
        const string _ATTENDANCE_STATUS_START_RANGE = "ATTENDANCE_STATUS!A";
        const string _ATTENDANCE_STATUS_END_RANGE = "D";
        const string _ATTENDANCE_STAT_ID_RANGE = "ATTENDANCE_STATUS!A:A";
        const string _ATTENDANCE_STAT_NAME_RANGE = "ATTENDANCE_STATUS!B:B";
        const string _LOG_RANGE = "LOG!A:C";

        const int _ATTENDANCE_STATUS_GID = 741019777;
        const int _ATTENDANCE_LOG_GID = 1617370344;


        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        SheetsService _service;
        UserCredential _credential;
        StringBuilder _exMsg;
        readonly string[] _scopes = { SheetsService.Scope.Spreadsheets };
        string _applicationName = "1732 Attendance Check-In Station";
        readonly string _sheetId = "13U-gYgtXlh8Q0Qgaim6nzrFlkOAP4dJP2hvOTaO7nTg";
        Dictionary<int, List<string>> dict_Attendance;

        #endregion

        #region *** PROPERTIES ***
        public string LastException { get { return _exMsg.ToString(); } }
        #endregion

        #region *** CONSTRUCTOR ***

        public GSheetsAPI()
        {
            _service = new SheetsService();
            _credential = null;
            dict_Attendance = new Dictionary<int, List<string>>();
        }

        #endregion

        #region *** FEATURE FUNCTIONALITY METHODS ***

        public bool Check_Valid_ID(int ID)
        {
            bool success = false;
            try
            {
                success = dict_Attendance.ContainsKey(ID);
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return success;
        }

        public string Check_ID_Status(int ID)
        {
            string status = string.Empty;
            try
            {
                status = dict_Attendance[ID][1];
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return status;
        }

        public string Get_ID_Name(int ID)
        {
            string name = string.Empty;
            try
            {
                name = dict_Attendance[ID][0];
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return name;
        }

        public bool Add_User(int ID, string name)
        {
            bool
                success = false;

            try
            {
                InsertRows(Create_Attendance_Status_Row(ID, name, "OUT"), string.Format("{0}{1}:{2}", _ATTENDANCE_STATUS_START_RANGE, Get_Next_Attendance_Row(), _ATTENDANCE_STATUS_END_RANGE));
                Read_Attendance_Status();
                InsertRows(Create_Log_Row(ID, _ADDED_STATUS), Get_Next_Log_Row());
                success = true;
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return success;
        }

        public bool Update_User_Status(int ID)
        {
            bool
                success = false;

            try
            {
                int rowToUpdate = Get_Attendance_Status_Row(ID);

                //row to be updated - increment by 1 because sheets start at "0"
                string rowRange = string.Format("{0}{1}:{2}{3}", _ATTENDANCE_STATUS_START_RANGE, (rowToUpdate + 1), _ATTENDANCE_STATUS_END_RANGE, (rowToUpdate + 1));

                UpdateRows(Create_Attendance_Status_Row(ID, dict_Attendance[ID][0], dict_Attendance[ID][1].Equals(_IN_STATUS) ? _OUT_STATUS : _IN_STATUS), rowRange);
                Read_Attendance_Status();
                InsertRows(Create_Log_Row(ID, dict_Attendance[ID][1]), Get_Next_Log_Row());
                success = true;
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return success;
        }

        public bool Delete_User(int ID)
        {
            bool
                success = false;

            try
            {
                int rowToRemove = Get_Attendance_Status_Row(ID);
                DeleteRows(rowToRemove);
                Read_Attendance_Status();
                InsertRows(Create_Log_Row(ID, _DELETED_STATUS), Get_Next_Log_Row());
                success = true;
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return success;
        }

        public bool Refresh_Local_Data()
        {
            bool
               success = false;

            try
            {
                Read_Attendance_Status();
                success = true;
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return success;
        }

        #endregion

        #region *** GET/READ METHODS ***
        private void Read_Attendance_Status()
        {
            try
            {
                GetRequest getRequest = _service.Spreadsheets.Values.Get(_sheetId, _ATTENDANCE_STATUS_RANGE);

                ValueRange getResponse = getRequest.Execute();
                IList<IList<Object>> idList = getResponse.Values;

                Parse_Attendance_Status_Rows(idList);
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        private int Get_Attendance_Status_Row(int ID)
        {
            int returnVal = -1;
            try
            {
                GetRequest getRequest = _service.Spreadsheets.Values.Get(_sheetId, _ATTENDANCE_STAT_ID_RANGE);

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

        private int Get_Next_Attendance_Row()
        {
            int returnVal = -1;
            try
            {
                GetRequest getRequest = _service.Spreadsheets.Values.Get(_sheetId, _ATTENDANCE_STATUS_RANGE);

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

        private string Get_Next_Log_Row()
        {
            string returnVal = string.Empty;
            try
            {
                GetRequest getRequest = _service.Spreadsheets.Values.Get(_sheetId, _LOG_RANGE);

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
        private void Parse_Attendance_Status_Rows(IList<IList<object>> idStatusList)
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

        private IList<IList<object>> Create_Log_Row(int ID, string status)
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

        private IList<IList<object>> Create_Attendance_Status_Row(int ID, string name, string status)
        {
            IList<IList<object>> newRow = new List<IList<object>>
            {
                new List<object>() { ID, name, status, DateTime.Now.ToString() }
            };
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

                    _credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        _scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                    Console.WriteLine("Credential file saved to: " + credPath);
                }

                // Create Google Sheets API service.
                _service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = _credential,
                    ApplicationName = _applicationName,
                });
                success = true;
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return success;
        }

        private void InsertRows(IList<IList<Object>> values, string newRange)
        {
            AppendRequest request = _service.Spreadsheets.Values.Append(new ValueRange() { Values = values }, _sheetId, newRange);
            request.InsertDataOption = AppendRequest.InsertDataOptionEnum.INSERTROWS;
            request.ValueInputOption = AppendRequest.ValueInputOptionEnum.RAW;
            AppendValuesResponse response = request.Execute();
        }

        private void UpdateRows(IList<IList<Object>> values, string newRange)
        {
            UpdateRequest request = _service.Spreadsheets.Values.Update(new ValueRange() { Values = values }, _sheetId, newRange);
            request.ValueInputOption = UpdateRequest.ValueInputOptionEnum.RAW;
            UpdateValuesResponse response = request.Execute();
        }

        private void DeleteRows(int rowToDelete)
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

            SpreadsheetsResource.BatchUpdateRequest Deletion = new SpreadsheetsResource.BatchUpdateRequest(_service, content, _sheetId);
            BatchUpdateSpreadsheetResponse response = Deletion.Execute();
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

            Console.Write(_exMsg.ToString());
        }
        #endregion
    }
}
