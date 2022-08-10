using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace SheetsQuickstart
{
    // Class to demonstrate the use of Sheets list values API
    class Program
    {
        /* Global instance of the scopes required by this quickstart.
         If modifying these scopes, delete your previously saved token.json/ folder. */
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Hoge App";

        static void Main(string[] args)
        {
            try
            {
                UserCredential credential;
                // Load client secrets.
                using (var stream =
                       new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                {
                    /* The file token.json stores the user's access and refresh tokens, and is created
                     automatically when the authorization flow completes for the first time. */
                    string credPath = AppDomain.CurrentDomain.BaseDirectory;
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.FromStream(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                    Console.WriteLine("Credentialの保存先: " + credPath);
                }

                CreateSheet();

                Console.Write("北(西野)：");
                var norceCountInNishino =  Console.ReadLine();
                Console.Write("北(ごっち)：");
                var norceCountInGotti = Console.ReadLine();
                Console.Write("北(なる)：");
                var norceCountInNaru = Console.ReadLine();
                Console.Write("北(はっしー)：");
                var norceCountInHasshi = Console.ReadLine();


                // Create Google Sheets API service.
                var service = new SheetsService(new BaseClientService.Initializer
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName
                });

                // Define request parameters.
                String spreadsheetId = "13D8Nzqr5STk14z6IY0td75uSgfkPeMBO6NnB_tE5VgQ";
                //String range = "Class Data!A2:E";
                //SpreadsheetsResource.ValuesResource.GetRequest request =
                //    service.Spreadsheets.Values.Get(spreadsheetId, range);

                var wv = new List<IList<object>>()
                {
                    new List<object>{norceCountInNishino, norceCountInGotti, norceCountInNaru, norceCountInHasshi}
                };
                var body = new ValueRange() { Values = wv };
                var request = service.Spreadsheets.Values.Append(body, spreadsheetId, "Sheet1!C3");
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;

                // Prints the names and majors of students in a sample spreadsheet:
                // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
                var response = request.Execute();

                Console.WriteLine("書き込み完了");
                Console.Read();
            }
            catch (FileNotFoundException e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private static CreateSheet()
        {

        }
    }
}