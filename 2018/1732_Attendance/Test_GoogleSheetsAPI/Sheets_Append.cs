using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Newtonsoft.Json;

namespace Test_GoogleSheetsAPI
{
    class Sheets_Append
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "TimeSheetUpdation By Cybria Technology";
        static string SheetId = "13U-gYgtXlh8Q0Qgaim6nzrFlkOAP4dJP2hvOTaO7nTg";

        static void Main(string[] args)
        {
            var service = AuthorizeGoogleApp();
            string nextRange = GetRange(service);
            IList<IList<object>> appRows = GenerateData();
            UpdatGoogleSheetinBatch(appRows, SheetId, nextRange, service);
        }

        private static IList<IList<object>> CreateRecord()
        {
            IList<IList<object>> appRows = new List<IList<object>>();
            IList<object> vals = new List<object>() { "2222", DateTime.UtcNow.ToString(), "IN" };
            appRows.Add(vals);
            return appRows;
        }

        private static SheetsService AuthorizeGoogleApp()
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
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            return service;
        }

        protected static string GetRange(SheetsService service)
        {
            // Define request parameters.
            String spreadsheetId = SheetId;
            String range = "ATTENDANCE!A:C";

            SpreadsheetsResource.ValuesResource.GetRequest getRequest =
                       service.Spreadsheets.Values.Get(spreadsheetId, range);

            ValueRange getResponse = getRequest.Execute();
            IList<IList<Object>> getValues = getResponse.Values;

            int currentCount = getValues.Count + 1;

            String newRange = "A" + currentCount + ":C";

            return newRange;
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
                string randDate = string.Format("{0}/{1}/{2} {3}:{4}:{5} {6}", new Random().Next(1, 12), new Random().Next(
                    1, 28), "2018", new Random().Next(0, 12), new Random().Next(0, 59), new Random().Next(0, 59), period);
                obj.Add(new Random().Next(1000, 99999));
                obj.Add(randDate);

                obj.Add(TS);
                objNewRecords.Add(obj);
            }

            return objNewRecords;
        }

        private static void UpdatGoogleSheetinBatch(IList<IList<Object>> values, string spreadsheetId, string newRange, SheetsService service)
        {
            SpreadsheetsResource.ValuesResource.AppendRequest request =
               service.Spreadsheets.Values.Append(new ValueRange() { Values = values }, spreadsheetId, newRange);
            request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
            var response = request.Execute();
        }
    }
}
