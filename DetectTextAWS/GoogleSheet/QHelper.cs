using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using System.IO;
using System.Threading;
using Google.Apis.Util.Store;
using System.Data;
using System.Windows.Forms;

namespace GoogleHelper
{
    class QHelper
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        string[] Scopes = { SheetsService.Scope.Spreadsheets };
        string ApplicationName = "GoogleSheetsHelper";
        SheetsService service;
        
        ~QHelper()
        {
            service.Dispose();
        }
        public SheetsService CreateService()
        {
            UserCredential credential;
            
            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
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

            return service;
        }

        public string CreateSheet(string sheetName)
        {
            var myNewSheet = new Google.Apis.Sheets.v4.Data.Spreadsheet();
            myNewSheet.Properties = new SpreadsheetProperties();
            myNewSheet.Properties.Title = sheetName;
            var newSpeadSheet = service.Spreadsheets.Create(myNewSheet).Execute();
            return newSpeadSheet.SpreadsheetId;
        }

        public void CreateEntry(String spreadsheetId, String range, IList<IList<object>> data)
        {
            var valueRange = new ValueRange();
            valueRange.Values = data;

            var appendRequest = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            appendRequest.Execute();
            data.Clear();
        }

        public IList<IList<Object>> ReadEntry(String spreadsheetId, String range)
        {
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);
            ValueRange response = request.Execute();
            return response.Values;
        }

        public void UpdateEntry(String spreadsheetId, String range, IList<IList<object>> data)
        {
            var valueRange = new ValueRange();
            valueRange.Values = data;

            var updateRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            updateRequest.Execute();
        }

        public void DeleteEntry(String SpreadsheetId, string range)
        {
            var requestBody = new ClearValuesRequest();

            var deleteRequest = service.Spreadsheets.Values.Clear(requestBody, SpreadsheetId, range);
            deleteRequest.Execute();
        }
    }
}
