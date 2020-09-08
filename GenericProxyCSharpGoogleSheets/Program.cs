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

namespace GoogleSheetsAPI4_v1console
{
    class Program
    {
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "GenericProxyCSharpGoogleSheets";
        static void Main(string[] args)
        {
            UserCredential credential;

            using (var stream = new FileStream("Secret/credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer(){ HttpClientInitializer = credential, ApplicationName = ApplicationName });

            // WIP Google Sheet
            // https://docs.google.com/spreadsheets/d/1sxsRFCFgfuxGWRFjZlt_twNZaNrVtysuFBkCNkQk1U4/edit#gid=0
            string spreadsheetId = "1sxsRFCFgfuxGWRFjZlt_twNZaNrVtysuFBkCNkQk1U4";
            string writeRange = "Ark1!A2:ZZ";

            var objList = new List<object>(){ DateTime.Now.ToLocalTime(), "Col2", "Col3", "Col4", "Col5", "My NEW Cell Text" };

            ValueRange valueDataRange = new ValueRange() { MajorDimension = "ROWS", Range = writeRange, Values = new List<IList<object>> { objList } };

            SpreadsheetsResource.ValuesResource.AppendRequest appendRequest = service.Spreadsheets.Values.Append(valueDataRange, spreadsheetId, writeRange);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
            appendRequest.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;

            AppendValuesResponse appendValueResponse = appendRequest.Execute();

            Console.WriteLine(JsonConvert.SerializeObject(valueDataRange));
            Console.ReadLine();
        }
    }
}