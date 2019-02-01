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

namespace _NET_1732_Attendance
{
    /// ***** PSC: Normal Use *****
    /// 1. Scan ID over sensor
    /// 2. Text populates on field - automatic carriage return will invoke keypress
    /// 3. Check for carriage return in keypress event and if present, continue check
    /// 4. Parse ID from field as ulong and check against existing dictionary read on startup of app (or from periodic invoke)
    /// 5. If ID is in dictionary, create new entry to record to LOG tab the ID and timestamp. 
    /// 5a. If ID is NOT in the directory, display a message to the screen.
    /// 6. Read the current status' of all IDs from the ATTENDANCE_STATUS tab
    /// 7. Enumerate current status of all IDs ulong dict_ID_Status
    /// 8. Verify current status of the ID and invert it to write to the ATTENDANCE_STATUS tab
    internal class GSheetsAPI : IDisposable
    {
        #region *** FIELDS ***
        const string _ROWS = "ROWS";
        const string _ADDED_STATUS = "ADDED";
        const string _UPDATED_STATUS = "UPDATED";
        const string _UNREGISTERED_STATUS = "UNREGISTERED";
        const string _HOURS_CREDITED = "CREDITED";
        const string _HOURS_MISSING = "MISSING";
        const string _DELETED_STATUS = "DELETED";
        const string _OUT_STATUS = "OUT";
        const string _IN_STATUS = "IN";

        const string _ATTENDANCE_STATUS = "ATTENDANCE_STATUS!";

        const string _ID_COL = "A";
        const string _SECONDARY_ID_COL = "B";
        const string _NAME_COL = "C";
        const string _IS_MENTOR_COL = "D";
        const string _CURRENT_STATUS_COL = "E";
        const string _CHECKIN_COL = "F";
        const string _HOURS_COL = "G";
        const string _CHECKOUT_COL = "H";
        const string _TOTAL_HRS_COL = "I";
        const string _TOTAL_MISSED_HRS_COL = "J";

        const string _RESET_HOURS = "00:00:00";
        const string _RESET_TOTAL_HOURS = "0.00:00:00";
        const string _NIGHTLY_CHECKOUT_TIME = "21:00:00";

        const string _LOG_START_RANGE = "LOG!A";
        const string _LOG_END_RANGE = "C";
        const string _LOG_RANGE = _LOG_START_RANGE + ":" + _LOG_END_RANGE;

        const string _ACCUM_HOURS = "ACCUMULATED_HOURS!";
        const string _DATE_RANGE = "1";
        const string _ID_NAME_RANGE = "A:B";

        const int _RECENT_CHECKOUT_TIME = 15;
        const int _GID_ATTENDANCE_STATUS = 741019777;
        const int _GID_ACCUMULATED_HOURS = 1137211462;
        const int _GID_ATTENDANCE_LOG = 1617370344;

        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        SheetsService _service;
        UserCredential _credential;
        StringBuilder _exMsg;
        readonly string[] _scopes = { SheetsService.Scope.Spreadsheets };
        string _applicationName = "1732 Attendance Check-In Station";
        Dictionary<ulong, User> dict_Attendance;
        Dictionary<ulong, ulong> dict_IDs;

        enum COLUMNS
        {
            ID = 0,
            SECONDARY_ID = 1,
            NAME = 2,
            IS_MENTOR = 3,
            STATUS = 4,
            LAST_CHECKIN = 5,
            HOURS = 6,
            LAST_CHECKOUT = 7,
            TOTAL_HOURS = 8,
            TOTAL_MISSED_HOURS = 9,
        }

        #endregion

        #region *** PROPERTIES ***        
        public string Sheet_ID { get; set; }
        public int GID_Attendance_Status { get; set; }
        public int GID_Accumulated_Hours { get; set; }
        public int GID_Attendance_Log { get; set; }
        public bool Auto_Checkout_Enabled { get; set; }
        public DateTime Auto_Checkout_Time { get; set; }
        public DateTime Team_Checkout_Time { get; set; }
        public DateTime Last_Data_Update { get; set; }
        public int Recent_Time_Check { get; set; }

        public string LastException { get { return _exMsg.ToString(); } }
        #endregion

        #region *** CONSTRUCTOR ***
        public GSheetsAPI()
        {
            _service = new SheetsService();
            _credential = null;
            dict_Attendance = new Dictionary<ulong, User>();
            dict_IDs = new Dictionary<ulong, ulong>();
        }

        public void Dispose()
        {
            ((IDisposable)_service).Dispose();
        }
        #endregion

        #region *** FEATURE FUNCTIONALITY METHODS ***

        public bool Check_Valid_ID(ulong ID, out ulong primaryID)
        {
            bool
                idFound = false;

            primaryID = 0;
            if (dict_Attendance.ContainsKey(ID))
            {
                primaryID = ID;
                idFound = true;
            }
            else if (dict_IDs.ContainsKey(ID))
            {
                primaryID = dict_IDs[ID];
                idFound = true;
            }

            return idFound;
        }

        public bool Check_Is_Mentor(ulong ID)
        {
            return dict_Attendance[ID].Is_Mentor;
        }

        public string Check_ID_Status(ulong ID)
        {
            return dict_Attendance[ID].Status;
        }

        public string Get_ID_Name(ulong ID)
        {
            return dict_Attendance[ID].Name;
        }

        public bool Add_User(ulong ID, ulong secondaryID, string name, ulong mentorID, bool isMentor)
        {
            bool
                success = false;

            try
            {
                InsertRows(Create_Attendance_Status_Row(ID, secondaryID, name, isMentor ? "X" : "", "OUT", "0", _RESET_HOURS, "0", _RESET_TOTAL_HOURS, _RESET_TOTAL_HOURS), string.Format("{0}{1}{2}:{3}", _ATTENDANCE_STATUS, _ID_COL, Get_Next_Attendance_Row(), _TOTAL_MISSED_HRS_COL));
                InsertRows(Create_Accumulated_Hours_Row(ID, name), string.Format("{0}{1}{2}:{3}", _ACCUM_HOURS, _ID_COL, Get_Next_Accumulated_Hours_Row(), _NAME_COL));
                Read_Attendance_Status();
                InsertRows(Create_Log_Row(ID, _ADDED_STATUS, mentorID), Get_Next_Log_Row());
                success = true;
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return success;
        }

        public bool Update_User(ulong ID, ulong secondaryID, string name, ulong mentorID, bool isMentor)
        {
            bool
                success = false;

            string
                secID = string.Empty;

            try
            {
                if (!secondaryID.Equals(0))
                {
                    secID = secondaryID.ToString();
                }

                UpdateRows(Create_Updated_Attendance_Status_Row(ID, secID, name, isMentor ? "X" : ""), string.Format("{0}{1}{2}:{3}", _ATTENDANCE_STATUS, _ID_COL, Get_User_Attendance_Status_Row(ID) + 1, _IS_MENTOR_COL));
                UpdateRows(Create_Accumulated_Hours_Row(ID, name), string.Format("{0}{1}{2}:{3}", _ACCUM_HOURS, _ID_COL, Get_Accumulated_Hours_User_Row(ID) + 1, _NAME_COL));
                Read_Attendance_Status();
                InsertRows(Create_Log_Row(ID, _UPDATED_STATUS, mentorID), Get_Next_Log_Row());
                success = true;
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return success;
        }

        public bool Credit_User_Hours(ulong ID, string creditedHours, ulong mentorID)
        {
            bool
                success = false;

            int
                rowToUpdate;

            string
                dateColumn = string.Empty,
                rowRange = string.Empty;

            try
            {
                // 1. Parse the hours from the form 
                TimeSpan.TryParse(creditedHours, out TimeSpan extraHours);
                if (extraHours.Equals(TimeSpan.Zero))
                {
                    throw new Exception("Invalid credited hours value passed");
                }

                // 2. Add the parsed hours to the current total hours value for the user
                TimeSpan sessionHours = extraHours + dict_Attendance[ID].User_TotalHours;

                // 3. Define the range (single cell) to update (ATTENDANCE_STATUS!H#)
                rowToUpdate = Get_User_Attendance_Status_Row(ID);
                rowRange = string.Format("{0}{1}{2}", _ATTENDANCE_STATUS, _TOTAL_HRS_COL, rowToUpdate + 1);

                // 4. Write the updated value
                UpdateRows(Update_Hours(sessionHours), rowRange);

                // 5. Even though the hours are credited - update the accumulated hours sheet
                dateColumn = Get_Accumulated_Hours_Date_Column();                                           //get the current date column
                rowToUpdate = Get_Accumulated_Hours_User_Row(ID);                                           //get the row of the user on that sheet
                string userCell = string.Format("{0}{1}", dateColumn, rowToUpdate + 1);                     //define the range of the cell to update

                //in case the user has split hours (left and came back)
                //read what is in the cell and add to it
                TimeSpan accumHours = Read_User_Accumulated_Hours(userCell) + extraHours;                   //tally the accumulated hours for today (if any) and add the extra hours calculated above
                UpdateRows(Update_Hours(accumHours), string.Format("{0}{1}", _ACCUM_HOURS, userCell));      //Write the updated value to the accumulated hours sheet

                // 5. Update the dictionary
                Read_Attendance_Status();

                // 6. Update the log
                InsertRows(Create_Log_Row(ID, _HOURS_CREDITED, mentorID), Get_Next_Log_Row());
                success = true;
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return success;
        }

        public bool Add_Missed_Hours(ulong ID, string missedHours, ulong mentorID)
        {
            bool
                success = false;

            int
                rowToUpdate;

            string
                dateColumn = string.Empty,
                rowRange = string.Empty;

            try
            {
                // 1. Parse the hours from the form 
                TimeSpan.TryParse(missedHours, out TimeSpan extraHours);

                // 2. Add the parsed hours to the current total hours value for the user
                TimeSpan missed = extraHours + dict_Attendance[ID].User_TotalMissedHours;

                // 3. Define the range (single cell) to update (ATTENDANCE_STATUS!H#)
                rowToUpdate = Get_User_Attendance_Status_Row(ID);
                rowRange = string.Format("{0}{1}{2}", _ATTENDANCE_STATUS, _TOTAL_MISSED_HRS_COL, rowToUpdate + 1);

                // 4. Write the updated value
                UpdateRows(Update_Hours(missed), rowRange);

                // 5. Update the dictionary
                Read_Attendance_Status();

                // 6. Update the log
                InsertRows(Create_Log_Row(ID, _HOURS_MISSING, mentorID), Get_Next_Log_Row());
                success = true;
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return success;
        }

        public bool Log_Unregistered_User(ulong ID)
        {
            bool
               success = false;

            try
            {
                InsertRows(Create_Log_Row(ID, _UNREGISTERED_STATUS), Get_Next_Log_Row());
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

            int
               rowToUpdate;

            string
               dateColumn = string.Empty,
               rowRange = string.Empty;

            try
            {
                rowToUpdate = Get_User_Attendance_Status_Row(ID);

                if (dict_Attendance[ID].Status.Equals(_IN_STATUS))
                {
                    /// If user currently checked in
                    /// 1) Set the user status to OUT
                    /// 2) Set the checked-out field 
                    /// 3) Calculate and set the hours field for that user (AFTER SETTING CHECKOUT TIME)
                    /// 4) Add the hours field to the total hours field

                    //calculate the hours and total time
                    dict_Attendance[ID].Status = _OUT_STATUS;
                    dict_Attendance[ID].Check_Out_Time = DateTime.Now;
                    dict_Attendance[ID].Calculate_Session_Hours();

                    //row to be updated - increment by 1 because sheets start at "0"
                    rowRange = string.Format("{0}{1}{2}", _ATTENDANCE_STATUS, _CURRENT_STATUS_COL, (rowToUpdate + 1));
                    UpdateRows(Update_Attendance_Status(dict_Attendance[ID].Status), rowRange);

                    rowRange = string.Format("{0}{1}{2}:{3}{4}", _ATTENDANCE_STATUS, _HOURS_COL, (rowToUpdate + 1), _TOTAL_HRS_COL, (rowToUpdate + 1));
                    UpdateRows(Update_Attendance_CheckOut(dict_Attendance[ID]), rowRange);

                    ///if current date is listed as a column and then search if user is in the list
                    ///if date does not exist, add the column
                    ///if the user does not exist on the sheet, add the user and then update their hours for that day
                    dateColumn = Get_Accumulated_Hours_Date_Column();
                    rowToUpdate = Get_Accumulated_Hours_User_Row(ID);
                    string userCell = string.Format("{0}{1}", dateColumn, rowToUpdate + 1);

                    //in case the user has split hours (left and came back)
                    //read what is in the cell and add to it
                    TimeSpan sessionHours = Read_User_Accumulated_Hours(userCell) + dict_Attendance[ID].User_Hours;
                    UpdateRows(Update_Hours(sessionHours), string.Format("{0}{1}", _ACCUM_HOURS, userCell));
                }
                else
                {
                    /// If user currently checked out
                    /// 1) Set the user status to IN
                    /// 2) Set the user check-in time
                    /// 3) Set the hours field for that user to 0
                    /// leave the rest of the data alone
                    dict_Attendance[ID].Status = _IN_STATUS;
                    dict_Attendance[ID].Check_In_Time = DateTime.Now;
                    TimeSpan.TryParse(_RESET_HOURS, out TimeSpan reset);
                    dict_Attendance[ID].User_Hours = reset;

                    rowRange = _ATTENDANCE_STATUS + _CURRENT_STATUS_COL + (rowToUpdate + 1) + ":" + _HOURS_COL + (rowToUpdate + 1);
                    UpdateRows(Update_Attendance_CheckIn(dict_Attendance[ID]), rowRange);
                }

                Read_Attendance_Status();
                InsertRows(Create_Log_Row(ID, dict_Attendance[ID].Status), Get_Next_Log_Row());
                success = true;
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return success;
        }

        public bool Delete_User(ulong ID, ulong mentorID)
        {
            bool
                success = false;

            try
            {
                int rowToRemove = Get_User_Attendance_Status_Row(ID);
                if (!rowToRemove.Equals(-1))
                {
                    //Delete user from ATTENDANCE STATUS tab
                    DeleteRows(rowToRemove, _GID_ATTENDANCE_STATUS);
                    Read_Attendance_Status();

                    //Delete user from ACCUMULATED HOURS tab
                    rowToRemove = Get_Accumulated_Hours_User_Row(ID);
                    DeleteRows(rowToRemove, _GID_ACCUMULATED_HOURS);

                    InsertRows(Create_Log_Row(ID, _DELETED_STATUS, mentorID), Get_Next_Log_Row());
                    success = true;
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return success;
        }

        public bool Force_Logoff_User(ulong ID)
        {
            bool
               success = false;

            int
               rowToUpdate;

            string
               dateColumn = string.Empty,
               rowRange = string.Empty;

            try
            {
                rowToUpdate = Get_User_Attendance_Status_Row(ID);

                /// If user is still checked in at the cut-off time 
                /// 1) Set the user status to OUT
                /// --- Do not accumulate hours 
                /// --- All hours are forfeit for that day 
                /// --- Calculate number of missed hours and add to the Total Missed Hours Column
                /// --- Leave all other fields alone 
                dict_Attendance[ID].Status = _OUT_STATUS;
                //if the user did not check out - calculate the total hours they missed
                dict_Attendance[ID].Calculate_Missed_Hours(new DateTime(DateTime.Now.Year, DateTime.Now.Month, dict_Attendance[ID].Check_In_Time.Day, Team_Checkout_Time.TimeOfDay.Hours, Team_Checkout_Time.TimeOfDay.Minutes, Team_Checkout_Time.TimeOfDay.Seconds));

                //row to be updated - increment by 1 because sheets start at "0"
                rowRange = string.Format("{0}{1}{2}", _ATTENDANCE_STATUS, _CURRENT_STATUS_COL, (rowToUpdate + 1));
                UpdateRows(Update_Attendance_Status(dict_Attendance[ID].Status), rowRange);

                //Update the missed hours cell
                UpdateRows(Update_Hours(dict_Attendance[ID].User_TotalMissedHours), string.Format("{0}{1}", _TOTAL_MISSED_HRS_COL, (rowToUpdate + 1)));

                ///update the dictionary?
                Read_Attendance_Status();
                InsertRows(Create_Log_Row(ID, dict_Attendance[ID].Status, 0, true), Get_Next_Log_Row());
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

        public List<User> Get_All_Users()
        {
            List<User> users = new List<User>();
            try
            {
                List<ulong> keys = new List<ulong>(dict_Attendance.Keys);
                foreach (ulong ID in keys)
                {
                    users.Add(dict_Attendance[ID]);
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return users;
        }

        public List<User> Get_CheckedIn_Users()
        {
            List<User> users = new List<User>();
            try
            {
                List<ulong> keys = new List<ulong>(dict_Attendance.Keys);
                foreach (ulong ID in keys)
                {
                    if (dict_Attendance[ID].Status.Equals(_IN_STATUS))
                    {
                        users.Add(dict_Attendance[ID]);
                    }
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return users;
        }

        public List<User> Get_Recently_CheckedOut_Users()
        {
            List<User> users = new List<User>();
            try
            {
                List<ulong> keys = new List<ulong>(dict_Attendance.Keys);
                foreach (ulong ID in keys)
                {
                    if ((DateTime.Now - dict_Attendance[ID].Check_Out_Time).TotalMinutes < Recent_Time_Check)
                    {
                        users.Add(dict_Attendance[ID]);
                    }
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return users;
        }

        public List<User> Get_Recently_CheckedIn_Users()
        {
            List<User> users = new List<User>();
            try
            {
                List<ulong> keys = new List<ulong>(dict_Attendance.Keys);
                foreach (ulong ID in keys)
                {
                    if ((DateTime.Now - dict_Attendance[ID].Check_In_Time).TotalMinutes < Recent_Time_Check)
                    {
                        users.Add(dict_Attendance[ID]);
                    }
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return users;
        }

        public List<string> Get_Full_Log()
        {
            IList<IList<object>> logRead;
            List<string> logEntries = new List<string>();
            string line = string.Empty;
            try
            {
                logRead = Get_Log_Entries();
                for (int i = 0; i < logRead.Count; i++)
                {
                    //Get the current row (ID | Current_Status)
                    IList<object> row = logRead[i];
                    foreach (string item in row)
                    {
                        line = line + string.Concat(item, "\t");
                    }
                    logEntries.Add(line);
                    line = string.Empty;
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return logEntries;
        }

        public List<string> Get_Log_100_Rows()
        {
            IList<IList<object>> logRead;
            List<string> logEntries = new List<string>();
            string line = string.Empty;
            try
            {
                logRead = Get_Log_Entries();
                if (logRead.Count > 100)
                {
                    int total = logRead.Count - 100;
                    for (int i = logRead.Count - 1; i != total; i--)
                    {
                        //Get the current row (ID | Current_Status)
                        IList<object> row = logRead[i];
                        foreach (string item in row)
                        {
                            line = line + string.Concat(item, "\t");
                        }
                        logEntries.Add(line);
                        line = string.Empty;
                    }
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return logEntries;
        }

        #endregion

        #region *** GET/READ METHODS ***
        private void Read_Attendance_Status()
        {
            try
            {
                GetRequest getRequest = _service.Spreadsheets.Values.Get(Sheet_ID, string.Format("{0}{1}:{2}", _ATTENDANCE_STATUS, _ID_COL, _TOTAL_MISSED_HRS_COL));

                ValueRange getResponse = getRequest.Execute();
                IList<IList<Object>> idList = getResponse.Values;

                Parse_Attendance_Status_Rows(idList);
            }
            catch (Exception ex)
            {
                throw new Exception(MethodBase.GetCurrentMethod().Name, ex);
            }
        }

        private TimeSpan Read_User_Accumulated_Hours(string userCell)
        {
            TimeSpan hoursRead;
            try
            {
                GetRequest getRequest = _service.Spreadsheets.Values.Get(Sheet_ID, string.Format("{0}{1}", _ACCUM_HOURS, userCell));

                ValueRange getResponse = getRequest.Execute();
                IList<IList<Object>> userHours = getResponse.Values;
                if (userHours != null)
                {
                    IList<object> cell = userHours[0];
                    TimeSpan.TryParse(cell[0].ToString(), out hoursRead);
                }
                else
                {
                    TimeSpan.TryParse(_RESET_TOTAL_HOURS, out hoursRead);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(MethodBase.GetCurrentMethod().Name, ex);
            }
            return hoursRead;
        }

        private int Get_User_Attendance_Status_Row(ulong ID)
        {
            int returnVal = -1;
            try
            {
                GetRequest getRequest = _service.Spreadsheets.Values.Get(Sheet_ID, string.Format("{0}{1}:{2}", _ATTENDANCE_STATUS, _ID_COL, _ID_COL));

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

        private int Get_Accumulated_Hours_User_Row(ulong ID)
        {
            int returnVal = -1;
            try
            {
                GetRequest getRequest = _service.Spreadsheets.Values.Get(Sheet_ID, string.Format("{0}{1}:{2}", _ACCUM_HOURS, _ID_COL, _ID_COL));

                ValueRange getResponse = getRequest.Execute();
                IList<IList<Object>> getValues = getResponse.Values;

                if (getValues != null)
                {
                    for (int i = 1; i < getValues.Count; i++)
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
                GetRequest getRequest = _service.Spreadsheets.Values.Get(Sheet_ID, string.Format("{0}{1}:{2}", _ATTENDANCE_STATUS, _ID_COL, _ID_COL));

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
                GetRequest getRequest = _service.Spreadsheets.Values.Get(Sheet_ID, string.Format("{0}{1}", _ACCUM_HOURS, _ID_NAME_RANGE));

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

        private string Get_Accumulated_Hours_Date_Column()
        {
            int
                lastColumn = -1;

            string
                returnVal = string.Empty;

            try
            {
                GetRequest getRequest = _service.Spreadsheets.Values.Get(Sheet_ID, string.Format("{0}{1}:{2}", _ACCUM_HOURS, _DATE_RANGE, _DATE_RANGE));

                ValueRange getResponse = getRequest.Execute();
                IList<IList<Object>> getValues = getResponse.Values;

                //if date column does not exist, create it
                //and return the index of the new column
                //otherwise return the index of the column
                if (!Parse_Accumulated_Hours_Dates(getValues, out lastColumn))
                {
                    InsertRows_Overwrite(Create_Accumulated_Hours_Date_Cell(), string.Format("{0}{1}{2}", _ACCUM_HOURS, Get_Next_Accum_Cell(lastColumn + 1), _DATE_RANGE));
                    returnVal = Get_Next_Accum_Cell(lastColumn + 1);
                }
                else
                {
                    returnVal = Get_Next_Accum_Cell(lastColumn);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(MethodBase.GetCurrentMethod().Name, ex);
            }
            return returnVal;
        }

        private IList<IList<object>> Get_Log_Entries()
        {
            try
            {
                GetRequest getRequest = _service.Spreadsheets.Values.Get(Sheet_ID, _LOG_RANGE);

                ValueRange getResponse = getRequest.Execute();
                IList<IList<Object>> getValues = getResponse.Values;
                return getValues;
            }
            catch (Exception ex)
            {
                throw new Exception(MethodBase.GetCurrentMethod().Name, ex);
            }
        }

        private string Get_Next_Log_Row()
        {
            string returnVal = string.Empty;
            try
            {
                GetRequest getRequest = _service.Spreadsheets.Values.Get(Sheet_ID, _LOG_RANGE);

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

        private string Get_Next_Accum_Cell(int index)
        {
            string
                columnLetter = string.Empty;

            int
                modulo,
                dividend = index;

            try
            {
                while (dividend > 0)
                {
                    modulo = (dividend - 1) % 26;
                    columnLetter = Convert.ToChar(65 + modulo).ToString() + columnLetter;
                    dividend = (dividend - modulo) / 26;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(MethodBase.GetCurrentMethod().Name, ex);
            }
            return columnLetter;
        }

        #endregion

        #region *** PARSE METHODS ***
        private void Parse_Attendance_Status_Rows(IList<IList<object>> idStatusList)
        {
            ulong
                ID,
                sec_ID = 0;

            try
            {
                //Wipe any previous dictionary values to start fresh with every request
                //Treats the Google Sheet as the golden copy
                dict_Attendance = new Dictionary<ulong, User>();
                dict_IDs = new Dictionary<ulong, ulong>();

                //Start at 1 because first row is the header row (ID | Name)
                for (int i = 1; i < idStatusList.Count; i++)
                {
                    //Get the current row (ID | Current_Status)
                    IList<object> row = idStatusList[i];
                    ulong.TryParse((string)row[(int)COLUMNS.ID], out ID);

                    if (!string.IsNullOrEmpty((string)row[(int)COLUMNS.SECONDARY_ID]))
                    {
                        ulong.TryParse((string)row[(int)COLUMNS.SECONDARY_ID], out sec_ID);
                    }

                    string name = row[(int)COLUMNS.NAME].ToString();
                    bool mentor = row[(int)COLUMNS.IS_MENTOR].ToString().Equals("X") ? true : false;
                    string stat = row[(int)COLUMNS.STATUS].ToString();
                    DateTime.TryParse(row[(int)COLUMNS.LAST_CHECKIN].ToString(), out DateTime lastCheckIn);
                    TimeSpan.TryParse(row[(int)COLUMNS.HOURS].ToString(), out TimeSpan hours);
                    DateTime.TryParse(row[(int)COLUMNS.LAST_CHECKOUT].ToString(), out DateTime lastCheckOut);
                    TimeSpan.TryParse(row[(int)COLUMNS.TOTAL_HOURS].ToString(), out TimeSpan totalHours);
                    TimeSpan.TryParse(row[(int)COLUMNS.TOTAL_MISSED_HOURS].ToString(), out TimeSpan totalMissedHours);

                    dict_Attendance.Add(ID, new User() { ID = ID, Name = name, Is_Mentor = mentor, Status = stat, Check_In_Time = lastCheckIn, User_Hours = hours, Check_Out_Time = lastCheckOut, User_TotalHours = totalHours, User_TotalMissedHours = totalMissedHours });

                    if (!string.IsNullOrEmpty((string)row[(int)COLUMNS.SECONDARY_ID]))
                    {
                        dict_IDs.Add(sec_ID, ID);
                        dict_Attendance[ID].Secondary_ID = sec_ID;
                    }
                }
                Last_Data_Update = DateTime.Now;
            }
            catch (Exception ex)
            {
                throw new Exception(MethodBase.GetCurrentMethod().Name, ex);
            }
        }

        private bool Parse_Accumulated_Hours_Dates(IList<IList<object>> dateRow, out int lastColumn)
        {
            bool success = false;
            try
            {
                IList<object> cell = dateRow[0];
                lastColumn = cell.Count;
                for (int i = cell.Count - 1; i > 0; i--)
                {
                    DateTime.TryParse(cell[i].ToString(), out DateTime date);
                    if (date.Equals(DateTime.Today))
                    {
                        success = true;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(MethodBase.GetCurrentMethod().Name, ex);
            }
            return success;

        }

        #endregion

        #region *** RECORD CREATION METHODS ***

        private IList<IList<object>> Create_Log_Row(ulong ID, string status, ulong mentorID = 0, bool forceLogOff = false)
        {
            string log = string.Empty;
            IList<IList<object>> newRow = new List<IList<object>>();

            switch (status)
            {
                case _IN_STATUS:
                    log = "User checked IN";
                    newRow.Add(new List<object>() { ID, DateTime.Now.ToString(), log });
                    break;
                case _OUT_STATUS:
                    log = forceLogOff ? "Forced user check OUT" : "User checked OUT";
                    newRow.Add(new List<object>() { ID, DateTime.Now.ToString(), log });
                    break;
                case _ADDED_STATUS:
                    log = string.Format("Added ID {0}", ID.ToString());
                    newRow.Add(new List<object>() { mentorID, DateTime.Now.ToString(), log });
                    break;
                case _UNREGISTERED_STATUS:
                    log = "User NOT REGISTERED";
                    newRow.Add(new List<object>() { ID, DateTime.Now.ToString(), log });
                    break;
                case _HOURS_CREDITED:
                    log = string.Format("Credited hours to ID {0}", ID.ToString());
                    newRow.Add(new List<object>() { mentorID, DateTime.Now.ToString(), log });
                    break;
                case _HOURS_MISSING:
                    log = string.Format("Added missed hours to ID {0}", ID.ToString());
                    newRow.Add(new List<object>() { mentorID, DateTime.Now.ToString(), log });
                    break;
                case _UPDATED_STATUS:
                    log = string.Format("Updated ID {0}", ID.ToString());
                    newRow.Add(new List<object>() { mentorID, DateTime.Now.ToString(), log });
                    break;
                case _DELETED_STATUS:
                    log = string.Format("Deleted ID {0}", ID.ToString());
                    newRow.Add(new List<object>() { mentorID, DateTime.Now.ToString(), log });
                    break;
                default:
                    throw new Exception("Unable to parse status to create row for Log sheet");
            }
            return newRow;
        }

        private IList<IList<object>> Create_Attendance_Status_Row(ulong ID, ulong secondaryID, string name, string isMentor, string status, string lastCheckIn, string hours, string lastCheckout, string totalHours, string totalMissedHours)
        {
            IList<IList<object>> newRow = new List<IList<object>>
            {
                new List<object>() { ID, secondaryID, name, isMentor, status, lastCheckIn, hours, lastCheckout, totalHours, totalMissedHours }
            };
            return newRow;
        }

        private IList<IList<object>> Create_Updated_Attendance_Status_Row(ulong ID, string secondaryID, string name, string isMentor)
        {
            IList<IList<object>> newRow = new List<IList<object>>
            {
                new List<object>() { ID, secondaryID, name, isMentor }
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

        private IList<IList<object>> Create_Accumulated_Hours_Date_Cell()
        {
            IList<IList<object>> newRow = new List<IList<object>>
            {
                new List<object>() { DateTime.Now.ToShortDateString() }
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

        private void InsertRows_Overwrite(IList<IList<Object>> values, string newRange)
        {
            AppendRequest request = _service.Spreadsheets.Values.Append(new ValueRange() { Values = values }, Sheet_ID, newRange);
            request.InsertDataOption = AppendRequest.InsertDataOptionEnum.OVERWRITE;
            request.ValueInputOption = AppendRequest.ValueInputOptionEnum.RAW;
            AppendValuesResponse response = request.Execute();
        }

        private void InsertRows(IList<IList<Object>> values, string newRange)
        {
            AppendRequest request = _service.Spreadsheets.Values.Append(new ValueRange() { Values = values }, Sheet_ID, newRange);
            request.InsertDataOption = AppendRequest.InsertDataOptionEnum.INSERTROWS;
            request.ValueInputOption = AppendRequest.ValueInputOptionEnum.RAW;
            AppendValuesResponse response = request.Execute();
        }

        private void UpdateRows(IList<IList<Object>> values, string newRange)
        {
            UpdateRequest request = _service.Spreadsheets.Values.Update(new ValueRange() { Values = values }, Sheet_ID, newRange);
            request.ValueInputOption = UpdateRequest.ValueInputOptionEnum.RAW;
            UpdateValuesResponse response = request.Execute();
        }

        private void DeleteRows(int rowToDelete, int sheetGID)
        {
            BatchUpdateSpreadsheetRequest content = new BatchUpdateSpreadsheetRequest();
            Request request = new Request()
            {
                DeleteDimension = new DeleteDimensionRequest()
                {
                    Range = new DimensionRange()
                    {
                        SheetId = sheetGID,
                        Dimension = _ROWS,
                        StartIndex = rowToDelete,
                        EndIndex = rowToDelete + 1
                    }
                }
            };

            List<Request> listRequests = new List<Request> { request };
            content.Requests = listRequests;

            SpreadsheetsResource.BatchUpdateRequest Deletion = new SpreadsheetsResource.BatchUpdateRequest(_service, content, Sheet_ID);
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

        private IList<IList<object>> Update_Attendance_CheckIn(User user)
        {
            IList<IList<object>> newRow = new List<IList<object>>
            {
                new List<object>() { user.Status, user.Check_In_Time.ToString(), user.User_Hours }
            };
            return newRow;
        }

        private IList<IList<object>> Update_Attendance_CheckOut(User user)
        {
            IList<IList<object>> newRow = new List<IList<object>>
            {
                new List<object>() { string.Format("{0:00}:{1:00}:{2:00}", user.User_Hours.Hours, user.User_Hours.Minutes, user.User_Hours.Seconds), user.Check_Out_Time.ToString(), string.Format("{0:00}.{1:00}:{2:00}:{3:00}", user.User_TotalHours.Days, user.User_TotalHours.Hours, user.User_TotalHours.Minutes, user.User_TotalHours.Seconds) }
            };
            return newRow;
        }

        private IList<IList<object>> Update_Hours(TimeSpan userSession)
        {
            IList<IList<object>> newRow = new List<IList<object>>
            {
                new List<object>() { string.Format("{0:00}.{1:00}:{2:00}:{3:00}", userSession.Days, userSession.Hours, userSession.Minutes, userSession.Seconds) }
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
