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
        private static SheetsService service;

        static void Main(string[] args)
        {
            // Google sheet used for storing configuration etc.
            // https://docs.google.com/spreadsheets/d/1sxsRFCFgfuxGWRFjZlt_twNZaNrVtysuFBkCNkQk1U4/edit#gid=0
            string rootSpreadsheetId = "1sxsRFCFgfuxGWRFjZlt_twNZaNrVtysuFBkCNkQk1U4";
            string rootSpreadsheetReadRange = "config!A2:ZZ";
            UserCredential credential;

            using (var stream = new FileStream("Secret/credentials.json", FileMode.Open, FileAccess.Read))
            {
                //string credPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
                string credPath = System.Environment.CurrentDirectory;

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            service = new SheetsService(new BaseClientService.Initializer(){ HttpClientInitializer = credential, ApplicationName = ApplicationName });

            // read configurations
            ValueRange configurations = Read(rootSpreadsheetId, rootSpreadsheetReadRange);
            Console.WriteLine(JsonConvert.SerializeObject(configurations));

            foreach (var configuration in configurations.Values)
            {
                string documentAlias = configuration[0].ToString();
                string documentID = configuration[1].ToString();
                string sheetName = configuration[2].ToString();

                Console.WriteLine($"document alias: {documentAlias},  document ID: {documentID}, sheet name: {sheetName}");
            }

            List<Object> objList = new List<object>() { DateTime.Now.ToLocalTime(), "My NEW Cell Text" };

            // Append row to sheet using Create method
            var appendValueResponse = Create(configurations.Values[1][1].ToString(), configurations.Values[1][2].ToString(), objList);

            Console.WriteLine(JsonConvert.SerializeObject(appendValueResponse));
            Console.ReadLine();
        }

        // CRUD methods (WIP)

        // Append row to sheet
        private static AppendValuesResponse Create(string documentID, string sheetName, List<Object> objList)
        {
            string writeRange = sheetName + "!A2:ZZ";
            ValueRange valueDataRange = new ValueRange() { MajorDimension = "ROWS", Range = writeRange, Values = new List<IList<object>> { objList }};
            SpreadsheetsResource.ValuesResource.AppendRequest appendRequest = service.Spreadsheets.Values.Append(valueDataRange, documentID, writeRange);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
            appendRequest.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
            return appendRequest.Execute();
        }
        private static ValueRange Read(string rootSpreadsheetId, string rootSpreadsheetReadRange)
        {
            SpreadsheetsResource.ValuesResource.GetRequest configRequest = service.Spreadsheets.Values.Get(rootSpreadsheetId, rootSpreadsheetReadRange);
            return configRequest.Execute();
        }
        private static void Update()
        {
            // Update
        }
        private static void Delete()
        {
            // Delete
        }
    }
}