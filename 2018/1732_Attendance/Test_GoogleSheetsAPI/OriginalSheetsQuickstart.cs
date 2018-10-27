using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.IO;
using System.Threading;

namespace Test_GoogleSheetsAPI
{
    class OriginalSheetsQuickstart
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "1732 Attendance Check-In Station";

        static void Main(string[] args)
        {
            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = GetCredential(),
                ApplicationName = ApplicationName,
            });

            // The ID of the spreadsheet to read/update.
            string spreadsheetId = "13U-gYgtXlh8Q0Qgaim6nzrFlkOAP4dJP2hvOTaO7nTg";
            // The A1 notation of a range to search for a logical table of data.
            // Values will be appended after the last row of the table.
            string range = "ATTENDANCE!A2:C";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);

            // Prints the names and majors of students in a sample spreadsheet:
            // https://docs.google.com/spreadsheets/d/13U-gYgtXlh8Q0Qgaim6nzrFlkOAP4dJP2hvOTaO7nTg/edit
            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                Console.WriteLine("ID, Timestamp, Status");
                foreach (var row in values)
                {
                    // Print columns A and E, which correspond to indices 0 and 4.
                    Console.WriteLine("{0}, {1}, {2}", row[0], row[1], row[2]);
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }
            Console.Read();

            IList<IList<object>> appRows = new List<IList<object>>();
            IList<object> vals = new List<object>() { "2222", DateTime.Now.ToString("yyyy-mm-dd hh:mm:ss"), "IN" };
            appRows.Add(vals);

            // How the input data should be interpreted.
            SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum valueInputOption =
                SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;  // TODO: Update placeholder value.

            // How the input data should be inserted.
            SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum insertDataOption =
                SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;  // TODO: Update placeholder value.

            // TODO: Assign values to desired properties of `requestBody`:
            ValueRange appendRequestBody = new ValueRange();
            appendRequestBody.MajorDimension = "ROWS";
            appendRequestBody.Values = appRows;

            SpreadsheetsResource.ValuesResource.AppendRequest appRequest = service.Spreadsheets.Values.Append(appendRequestBody, spreadsheetId, range);
            appRequest.ValueInputOption = valueInputOption;
            appRequest.InsertDataOption = insertDataOption;

            // To execute asynchronously in an async method, replace `request.Execute()` as shown:
            AppendValuesResponse appResponse = appRequest.Execute();
            // Data.AppendValuesResponse response = await request.ExecuteAsync();

            // TODO: Change code below to process the `response` object:
            Console.WriteLine(JsonConvert.SerializeObject(response));
        }

        public static UserCredential GetCredential()
        {
            // TODO: Change placeholder below to generate authentication credentials. See:
            // https://developers.google.com/sheets/quickstart/dotnet#step_3_set_up_the_sample
            //
            // Authorize using one of the following scopes:
            //     "https://www.googleapis.com/auth/drive"
            //     "https://www.googleapis.com/auth/drive.file"
            //     "https://www.googleapis.com/auth/spreadsheets"
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                //string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None
                    //,new FileDataStore(credPath, true)
                    ).Result;
                //Console.WriteLine("Credential file saved to: " + credPath);
            }
            return credential;
        }
    }
}
