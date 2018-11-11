using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource;
using System.Text;
using System.Reflection;
using System.Diagnostics;

namespace Test_GoogleSheetsAPI
{
    class Sheets_Append
    {
        const string _ATTENDANCE_LOG_RANGE = "LOG!A:C";
        const string _ATTENDANCE_STATUS_RANGE = "ATTENDANCE_STATUS!A:B";
        const string _IDS_RANGE = "IDS!A:B";

        enum ATTENDANCE_STATUS
        {
            IN,
            OUT,
            UNKNOWN
        }


        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static SheetsService service;
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "1732 Attendance Check-In Station";
        static string SheetId = "13U-gYgtXlh8Q0Qgaim6nzrFlkOAP4dJP2hvOTaO7nTg";
        static Dictionary<int, string> dict_ID_Name = new Dictionary<int, string>();
        static Dictionary<int, ATTENDANCE_STATUS> dict_ID_Status = new Dictionary<int, ATTENDANCE_STATUS>();

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
            watch.Start();
            AuthorizeGoogleApp();
            watch.Stop();
            Console.WriteLine(string.Format("T2-AUTH: {0}ms", watch.ElapsedMilliseconds));

            watch.Restart();
            Get_Current_ID_List();
            watch.Stop();
            Console.WriteLine(string.Format("T2-ID|NAME: {0}ms", watch.ElapsedMilliseconds));

            watch.Restart();
            Get_Current_ID_Status();
            watch.Stop();
            Console.WriteLine(string.Format("T2-ID|STATUS: {0}ms", watch.ElapsedMilliseconds));

            int idResult;
            Console.WriteLine("Enter an ID to check for");
            int.TryParse(Console.ReadLine(), out idResult);
            CheckID(idResult);

            Console.ReadLine();
            //IList<IList<object>> appRows = GenerateData();
            //UpdateGoogleSheetinBatch(appRows, SheetId, nextRange, service);
        }



        #region *** GET METHODS ***
        protected static void Get_Current_ID_List()
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

        private static void Get_Current_ID_Status()
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

        protected static string Get_NextAttendanceRow()
        {
            string returnVal = string.Empty;
            try
            {
                GetRequest getRequest = service.Spreadsheets.Values.Get(SheetId, _ATTENDANCE_LOG_RANGE);

                ValueRange getResponse = getRequest.Execute();
                IList<IList<Object>> getValues = getResponse.Values;

                returnVal = string.Format("A{0}:C", getValues.Count + 1);
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
            return returnVal;
        }
        #endregion

        #region *** CHECK/PARSE METHODS ***
        private static void CheckID(int checkID)
        {
            try
            {
                if (dict_ID_Name.ContainsKey(checkID) && dict_ID_Status.ContainsKey(checkID))
                {
                    Console.WriteLine("Exists!");
                    Console.WriteLine("Name: " + dict_ID_Name[checkID]);
                    Console.WriteLine("Status: " + dict_ID_Status[checkID]);
                }
                else if (dict_ID_Name.ContainsKey(checkID))
                {
                    Console.WriteLine("Only exists in ID|Name list!");
                    Console.WriteLine("Name: " + dict_ID_Name[checkID]);

                    string read;
                    do
                    {
                        Console.WriteLine("Add to ID|STATUS List?");
                        read = Console.ReadLine().Trim().ToUpper();
                    }
                    while (!read.Equals("Y") || !read.Equals("N"));

                    if(read.Equals("Y"))
                        InsertRows(Insert_ID_Status_Row(checkID, ATTENDANCE_STATUS.OUT), SheetId, _ATTENDANCE_STATUS_RANGE, service);

                }
                else if (dict_ID_Status.ContainsKey(checkID))
                {
                    Console.WriteLine("Only exists in ID|STATUS list!");
                    Console.WriteLine("Status: " + dict_ID_Status[checkID]);
                    string read;
                    do
                    {
                        Console.WriteLine("Add to ID|Name List?");
                        read = Console.ReadLine().Trim().ToUpper();
                    }
                    while (!read.Equals("Y") || !read.Equals("N"));

                    if (read.Equals("Y"))
                    {
                        Console.WriteLine("Please enter a name:");
                        read = Console.ReadLine();
                        InsertRows(Insert_ID_Name_Row(checkID, read), SheetId, _IDS_RANGE, service);

                    }
                }
                else
                    Console.WriteLine("Didn't find that ID in any list");
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        private static void Parse_IDS_Name(IList<IList<object>> idNameList)
        {
            try
            {
                //Start at 1 because first row is the header (ID | Name)
                for (int i = 1; i < idNameList.Count; i++)
                {
                    IList<object> row = idNameList[i];
                    int numId;
                    int.TryParse((string)row[0], out numId);
                    string name = (string)row[1];
                    dict_ID_Name.Add(numId, name);
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        private static void Parse_IDS_Status(IList<IList<object>> idStatusList)
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

                    dict_ID_Status.Add(numId, stat);
                }
            }
            catch (Exception ex)
            {
                HandleException(ex, MethodBase.GetCurrentMethod().Name);
            }
        }

        #endregion

        #region *** RECORD CREATION METHODS ***

        private static IList<IList<object>> Insert_ID_Name_Row(int ID, string name)
        {
            IList<IList<object>> newRow = new List<IList<object>>();
            newRow.Add(new List<object>() { ID, name });
            return newRow;
        }

        private static IList<IList<object>> Insert_ID_Timestamp_Row(int ID)
        {
            IList<IList<object>> newRow = new List<IList<object>>();
            newRow.Add(new List<object>() { ID, DateTime.UtcNow.ToString() });
            return newRow;
        }

        private static IList<IList<object>> Insert_ID_Status_Row(int ID, ATTENDANCE_STATUS status)
        {
            IList<IList<object>> newRow = new List<IList<object>>();
            newRow.Add(new List<object>() { ID, status.ToString() });
            return newRow;
        }

        private static IList<IList<object>> CreateRecord()
        {
            IList<IList<object>> appRows = new List<IList<object>>();
            IList<object> vals = new List<object>() { "2222", DateTime.UtcNow.ToString(), "IN" };
            appRows.Add(vals);
            return appRows;
        }

        private static IList<IList<Object>> GenerateData()
        {
            List<IList<Object>> objNewRecords = new List<IList<Object>>();

            IList<Object> obj = new List<Object>();
            for (int i = 0; i < 10; i++)
            {
                obj = new List<Object>();
                int val = new Random().Next(0, 1);
                string period = val.Equals(0) ? "AM" : "PM";
                val = new Random().Next(0, 1);
                string TS = val.Equals(0) ? "IN" : "OUT";
                string randDate = string.Format("{0}/{1}/{2} {3}:{4}:{5} {6}",
                    new Random().Next(1, 12),
                    new Random().Next(1, 28),
                    "2018",
                    new Random().Next(0, 12),
                    new Random().Next(0, 59),
                    new Random().Next(0, 59),
                    period);
                obj.Add(new Random().Next(1000, 99999));
                obj.Add(randDate);

                obj.Add(TS);
                objNewRecords.Add(obj);
            }

            return objNewRecords;
        }
        #endregion

        #region *** CONNECT/SEND DATA METHODS ***
        private static void AuthorizeGoogleApp()
        {
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
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
        }

        private static void InsertRows(IList<IList<Object>> values, string spreadsheetId, string newRange, SheetsService service)
        {
            AppendRequest request = service.Spreadsheets.Values.Append(new ValueRange() { Values = values }, spreadsheetId, newRange);
            request.InsertDataOption = AppendRequest.InsertDataOptionEnum.INSERTROWS;
            request.ValueInputOption = AppendRequest.ValueInputOptionEnum.RAW;
            var response = request.Execute();
        }

        #endregion

        #region *** EXCEPTION/GUI HANDLING ***
        private static void HandleException(Exception ex, string callingMethod)
        {
            StringBuilder exMsg = new StringBuilder();

            exMsg.AppendLine(string.Format("Exception thrown in: {0}", callingMethod));
            exMsg.AppendLine(string.IsNullOrEmpty(ex.Message) ? "" : ex.Message);
            exMsg.AppendLine(string.IsNullOrEmpty(ex.Source) ? "" : ex.Source);
            exMsg.AppendLine(string.IsNullOrEmpty(ex.StackTrace.ToString()) ? "" : ex.StackTrace.ToString());

            Console.WriteLine("");
            Console.WriteLine(exMsg);
            Console.WriteLine("");
        }
        #endregion
    }
}
