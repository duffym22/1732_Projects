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
    /// 4. Parse ID from field as ulongeger and check against existing dictionary read on startup of app (or from periodic invoke)
    /// 5. If ID is in dictionary, create new entry to record to LOG tab the ID and timestamp. 
    /// 5a. If ID is NOT in the directory, display to screen "ID: [IDVAL] is not in the list of valid IDs. Please contact a mentor to be added"
    /// 6. Read the current status' of all IDs from the ATTENDANCE_STATUS tab
    /// 7. Enumerate current status of all IDs ulongo dict_ID_Status
    /// 8. Verify current status of the ID and invert it to write to the ATTENDANCE_STATUS tab
    class GSheetsAPI
    {
        #region *** FIELDS ***
        const string _ROWS = "ROWS";
        const string _ADDED_STATUS = "ADDED";
        const string _DELETED_STATUS = "DELETED";
        const string _OUT_STATUS = "OUT";
        const string _IN_STATUS = "IN";


        const string _ATTENDANCE_STATUS = "ATTENDANCE_STATUS!";

        const string _ID_COL = "A";
        const string _NAME_COL = "B";
        const string _CURRENT_STATUS_COL = "C";
        const string _CHECKIN_COL = "D";
        const string _HOURS_COL = "E";
        const string _CHECKOUT_COL = "F";
        const string _TOTAL_HRS_COL = "G";

        const string _RESET_HOURS = "00:00:00";
        const string _RESET_TOTAL_HOURS = "0.00:00:00";

        const string _LOG_START_RANGE = "LOG!A";
        const string _LOG_END_RANGE = "C";
        const string _LOG_RANGE = _LOG_START_RANGE + ":" + _LOG_END_RANGE;

        const string _ACCUM_HOURS_DATE_RANGE = "ACCUMULATED_HOURS!1:1";
        const string _ACCUM_HOURS_START_RANGE = "ACCUMULATED_HOURS!A";
        const string _ACCUM_HOURS_END_RANGE = "B";
        const string _ACCUM_HOURS_ID_NAME_RANGE = "ACCUMULATED_HOURS!A:B";
        const string _ACCUM_HOURS_TOTAL_RANGE = _ACCUM_HOURS_START_RANGE + ":" + _ACCUM_HOURS_END_RANGE;

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
        Dictionary<ulong, List<string>> dict_Attendance;

        enum COLUMNS
        {
            ID = 0,
            NAME = 1,
            STATUS = 2,
            LAST_CHECKIN = 3,
            HOURS = 4,
            LAST_CHECKOUT = 5,
            TOTAL_HOURS = 6,
        }

        #endregion

        #region *** PROPERTIES ***
        public string LastException { get { return _exMsg.ToString(); } }
        #endregion

        #region *** CONSTRUCTOR ***

        public GSheetsAPI()
        {
            _service = new SheetsService();
            _credential = null;
            dict_Attendance = new Dictionary<ulong, List<string>>();
        }

        #endregion

        #region *** FEATURE FUNCTIONALITY METHODS ***

        public bool Check_Valid_ID(ulong ID)
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

        public string Check_ID_Status(ulong ID)
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

        public string Get_ID_Name(ulong ID)
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

        public bool Add_User(ulong ID, string name)
        {
            bool
                success = false;

            try
            {
                InsertRows(Create_Attendance_Status_Row(ID, name, "OUT", "0", _RESET_HOURS, "0", _RESET_TOTAL_HOURS), string.Format("{0}{1}{2}:{3}", _ATTENDANCE_STATUS, _ID_COL, Get_Next_Attendance_Row(), _TOTAL_HRS_COL));
                InsertRows(Create_Accumulated_Hours_Row(ID, name), string.Format("{0}{1}:{2}", _ACCUM_HOURS_START_RANGE, Get_Next_Accumulated_Hours_Row(), _ACCUM_HOURS_END_RANGE));
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

        public bool Update_User_Status(ulong ID)
        {
            bool
                success = false;

            string
                rowRange = string.Empty;

            try
            {
                int rowToUpdate = Get_User_Attendance_Status_Row(ID);
                string status = dict_Attendance[ID][1].Equals(_IN_STATUS) ? _OUT_STATUS : _IN_STATUS;
                string lastCheckIn = dict_Attendance[ID][2];
                string hours = dict_Attendance[ID][3];
                string lastCheckOut = dict_Attendance[ID][4];
                string totalHours = dict_Attendance[ID][5];


                if (dict_Attendance[ID][1].Equals(_IN_STATUS))
                {
                    /// If user currently checked in
                    /// 1) Set the user status to OUT
                    /// 2) Calculate and set the hours field for that user
                    /// 3) Set the checked-out field 
                    /// 4) Add the hours field to the total hours field

                    //calculate the hours and total time
                    DateTime dLastCheckOut = DateTime.Now;
                    DateTime.TryParse(lastCheckIn, out DateTime dlastCheckIn);
                    TimeSpan.TryParse(hours, out TimeSpan hoursResult);
                    TimeSpan.TryParse(totalHours, out TimeSpan totalHoursResult);
                    TimeSpan timeSpan = dLastCheckOut - dlastCheckIn;

                    hoursResult += timeSpan;
                    totalHoursResult += timeSpan;

                    //row to be updated - increment by 1 because sheets start at "0"
                    rowRange = _ATTENDANCE_STATUS + _CURRENT_STATUS_COL + (rowToUpdate + 1);
                    UpdateRows(Update_Attendance_Status(status), rowRange);

                    rowRange = _ATTENDANCE_STATUS + _HOURS_COL + (rowToUpdate + 1) + ":" + _TOTAL_HRS_COL + (rowToUpdate + 1);
                    UpdateRows(Update_Attendance_CheckOut(
                        string.Format("{0:00}:{1:00}:{2:00}", hoursResult.TotalHours, hoursResult.TotalMinutes, hoursResult.TotalSeconds),
                        dLastCheckOut.ToString(),
                        string.Format("{0:00}.{1:00}:{2:00}:{3:00}", totalHoursResult.Days, totalHoursResult.Hours, totalHoursResult.Minutes, totalHoursResult.Seconds)),
                        rowRange);

                    ///TODO - Add UpdateRows(Update_Accumulated_Hours()) function to search 
                    ///
                    ///if current date is listed as a column and then search if user is in the list
                    ///if date does not exist, add the column
                    ///if the user does not exist on the sheet, add the user and then update their hours for that day
                }
                else
                {
                    ///If user currently checked out 
                    /// 1) Set the user status to IN
                    /// 2) Set the user check-in time
                    /// 3) Set the hours field for that user to 0
                    /// leave the rest of the data alone
                    rowRange = _ATTENDANCE_STATUS + _CURRENT_STATUS_COL + (rowToUpdate + 1) + ":" + _HOURS_COL + (rowToUpdate + 1);
                    UpdateRows(Update_Attendance_CheckIn(status, DateTime.Now.ToString(), _RESET_HOURS), rowRange);
                }

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

        public bool Delete_User(ulong ID)
        {
            bool
                success = false;

            try
            {
                int rowToRemove = Get_User_Attendance_Status_Row(ID);
                if (!rowToRemove.Equals(-1))
                {
                    DeleteRows(rowToRemove);
                    Read_Attendance_Status();
                    InsertRows(Create_Log_Row(ID, _DELETED_STATUS), Get_Next_Log_Row());
                    success = true;
                }
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
                GetRequest getRequest = _service.Spreadsheets.Values.Get(_sheetId, string.Format("{0}{1}:{2}", _ATTENDANCE_STATUS, _ID_COL, _TOTAL_HRS_COL));

                ValueRange getResponse = getRequest.Execute();
                IList<IList<Object>> idList = getResponse.Values;

                Parse_Attendance_Status_Rows(idList);
            }
            catch (Exception ex)
            {
                throw new Exception(MethodBase.GetCurrentMethod().Name, ex);
            }
        }

        private int Get_User_Attendance_Status_Row(ulong ID)
        {
            int returnVal = -1;
            try
            {
                GetRequest getRequest = _service.Spreadsheets.Values.Get(_sheetId, string.Format("{0}{1}:{2}", _ATTENDANCE_STATUS, _ID_COL, _ID_COL));

                ValueRange getResponse = getRequest.Execute();
                IList<IList<Object>> getValues = getResponse.Values;

                if (getValues != null)
                {
                    for (int i = 0; i < getValues.Count; i++)
                    {
                        IList<object> row = getValues[i];
                        ulong.TryParse(row[0].ToString(), out ulong readID);
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
                throw new Exception(MethodBase.GetCurrentMethod().Name, ex);
            }
            return returnVal;
        }

        private int Get_Next_Attendance_Row()
        {
            int returnVal = -1;
            try
            {
                GetRequest getRequest = _service.Spreadsheets.Values.Get(_sheetId, string.Format("{0}{1}:{2}", _ATTENDANCE_STATUS, _ID_COL, _ID_COL));

                ValueRange getResponse = getRequest.Execute();
                IList<IList<Object>> getValues = getResponse.Values;

                returnVal = getValues.Count + 1;
            }
            catch (Exception ex)
            {
                throw new Exception(MethodBase.GetCurrentMethod().Name, ex);
            }
            return returnVal;
        }

        private int Get_Next_Accumulated_Hours_Row()
        {
            int returnVal = -1;
            try
            {
                GetRequest getRequest = _service.Spreadsheets.Values.Get(_sheetId, _ACCUM_HOURS_ID_NAME_RANGE);

                ValueRange getResponse = getRequest.Execute();
                IList<IList<Object>> getValues = getResponse.Values;

                returnVal = getValues.Count + 1;
            }
            catch (Exception ex)
            {
                throw new Exception(MethodBase.GetCurrentMethod().Name, ex);
            }
            return returnVal;
        }

        private int Find_Date_Column_Accumulated_Hours()
        {
            int returnVal = -1;
            try
            {
                GetRequest getRequest = _service.Spreadsheets.Values.Get(_sheetId, _ACCUM_HOURS_DATE_RANGE);

                ValueRange getResponse = getRequest.Execute();
                IList<IList<Object>> getValues = getResponse.Values;

                returnVal = getValues.Count + 1;
            }
            catch (Exception ex)
            {
                throw new Exception(MethodBase.GetCurrentMethod().Name, ex);
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

                returnVal = string.Format("{0}{1}:{2}", _LOG_START_RANGE, getValues.Count + 1, _LOG_END_RANGE);
            }
            catch (Exception ex)
            {
                throw new Exception(MethodBase.GetCurrentMethod().Name, ex);
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
                dict_Attendance = new Dictionary<ulong, List<string>>();

                //Start at 1 because first row is the header row (ID | Name)
                for (int i = 1; i < idStatusList.Count; i++)
                {
                    //Get the current row (ID | Current_Status)
                    IList<object> row = idStatusList[i];
                    ulong.TryParse((string)row[(int)COLUMNS.ID], out ulong ID);
                    name = row[(int)COLUMNS.NAME].ToString();
                    stat = row[(int)COLUMNS.STATUS].ToString();
                    DateTime.TryParse(row[(int)COLUMNS.LAST_CHECKIN].ToString(), out DateTime lastCheckIn);
                    TimeSpan.TryParse(row[(int)COLUMNS.HOURS].ToString(), out TimeSpan hours);
                    DateTime.TryParse(row[(int)COLUMNS.LAST_CHECKOUT].ToString(), out DateTime lastCheckOut);
                    TimeSpan.TryParse(row[(int)COLUMNS.TOTAL_HOURS].ToString(), out TimeSpan totalHours);
                    dict_Attendance.Add(ID, new List<string> { name, stat, lastCheckIn.ToString(), hours.ToString(), lastCheckOut.ToString(), totalHours.ToString() });
                }
            }
            catch (Exception ex)
            {
                throw new Exception(MethodBase.GetCurrentMethod().Name, ex);
            }
        }

        #endregion

        #region *** RECORD CREATION METHODS ***

        private IList<IList<object>> Create_Log_Row(ulong ID, string status)
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

        private IList<IList<object>> Create_Attendance_Status_Row(ulong ID, string name, string status, string lastCheckIn, string hours, string lastCheckout, string totalHours)
        {
            IList<IList<object>> newRow = new List<IList<object>>
            {
                new List<object>() { ID, name, status, lastCheckIn, hours, lastCheckout, totalHours }
            };
            return newRow;
        }

        private IList<IList<object>> Create_Accumulated_Hours_Row(ulong ID, string name)
        {
            IList<IList<object>> newRow = new List<IList<object>>
            {
                new List<object>() { ID, name }
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
                throw new Exception(MethodBase.GetCurrentMethod().Name, ex);
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
                        Dimension = _ROWS,
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

        private IList<IList<object>> Update_Attendance_Status(string status)
        {
            IList<IList<object>> newRow = new List<IList<object>>
            {
                new List<object>() { status }
            };
            return newRow;
        }

        private IList<IList<object>> Update_Attendance_CheckIn(string status, string lastCheckIn, string hours)
        {
            IList<IList<object>> newRow = new List<IList<object>>
            {
                new List<object>() { status, lastCheckIn, hours }
            };
            return newRow;
        }

        private IList<IList<object>> Update_Attendance_CheckOut(string hours, string lastCheckOut, string totalHours)
        {
            IList<IList<object>> newRow = new List<IList<object>>
            {
                new List<object>() { hours, lastCheckOut, totalHours }
            };
            return newRow;
        }

        private IList<IList<object>> Update_Accumulated_Hours(string hours)
        {
            IList<IList<object>> newRow = new List<IList<object>>
            {
                new List<object>() { hours }
            };
            return newRow;
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
