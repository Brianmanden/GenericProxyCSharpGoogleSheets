//using Google.Apis.Auth.OAuth2;
//using Google.Apis.Sheets.v4;
//using Google.Apis.Sheets.v4.Data;
//using Google.Apis.Services;
//using Google.Apis.Util.Store;
//using System;
//using System.Collections.Generic;
//using System.IO;
//using System.Threading;

//namespace CSharpGoogleSheets
//{
//    class Program
//    {
//        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
//        static string ApplicationName = "GenericProxyCSharpGoogleSheets";

//        static void Main(string[] args)
//        {
//            UserCredential credential;

//            using (
//                var stream = new FileStream("Secret/credentials.json", FileMode.Open, FileAccess.Read))
//            {
//                // The file token.json stores the user´s access and refresh tokens, and is created
//                // automatically when the authorization flow completes for the first time.

//                string credPath = "token.json";
//                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
//                    GoogleClientSecrets.Load(stream).Secrets,
//                    Scopes,
//                    "user",
//                    CancellationToken.None,
//                    new FileDataStore(credPath, true)).Result;

//                Console.WriteLine("Credential file saved to: " + credPath);
//            }

//            // Create Google Sheets API service.
//            var service = new SheetsService(new BaseClientService.Initializer()
//            {
//                HttpClientInitializer = credential,
//                ApplicationName = ApplicationName
//            });

//            // Define request parameters.

//            // Student Name	Gender	Class Level	Home State	Major	Extracurricular Activity
//            IList<Object> obj = new List<Object>();
//            obj.Add("A2");
//            obj.Add("B2");
//            IList<IList<Object>> values = new List<IList<Object>>();
//            values.Add(obj);

//            ValueRange body = new ValueRange();
//            body.Values = values;

//            // https://docs.google.com/spreadsheets/d/1sxsRFCFgfuxGWRFjZlt_twNZaNrVtysuFBkCNkQk1U4/edit#gid=0
//            String spreadsheetId = "1sxsRFCFgfuxGWRFjZlt_twNZaNrVtysuFBkCNkQk1U4";

//            String range = "Class Data!A2:E";
//            //SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);

//            SpreadsheetsResource.ValuesResource.AppendRequest appendRequest = service.Spreadsheets.Values.Append(body, spreadsheetId, range);
//            appendRequest.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
//            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;

//            // Prints the names and majors of students in an example spreadsheet:
//            // Google Example Sheet
//            // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
//            // My working Sheet
//            // https://docs.google.com/spreadsheets/d/1sxsRFCFgfuxGWRFjZlt_twNZaNrVtysuFBkCNkQk1U4/edit#gid=0
//            AppendValuesResponse response = appendRequest.Execute();

//            //IList<IList<Object>> values = response.Values;
//            //if (values != null && values.Count > 0)
//            //{
//            //    Console.WriteLine("Name, Major");
//            //    foreach (var row in values)
//            //    {
//            //        // Print columns A and E, which correspond to indices 0 and 4.
//            //        Console.WriteLine("{0}, {1}", row[0], row[4]);
//            //    }
//            //}
//            //else
//            //{
//            //    Console.WriteLine("No data found.");
//            //}
//            Console.ReadLine();
//        }
//    }
//}

using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace GoogleSheetsAPI4_v1console
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.Spreadsheets }; // static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "GenericProxyCSharpGoogleSheets";

        static void Main(string[] args)
        {
            UserCredential credential;

            using (var stream = new FileStream("Secret/credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);

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



            String spreadsheetId2 = "1sxsRFCFgfuxGWRFjZlt_twNZaNrVtysuFBkCNkQk1U4";
            String range2 = "F5";  // update cell F5 
            ValueRange valueRange = new ValueRange();
            valueRange.MajorDimension = "COLUMNS";//"ROWS";//COLUMNS

            var oblist = new List<object>() { "My Cell Text" };
            valueRange.Values = new List<IList<object>> { oblist };

            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId2, range2);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            UpdateValuesResponse result2 = update.Execute();

            Console.WriteLine("done!");
            Console.ReadLine();
        }
    }
}