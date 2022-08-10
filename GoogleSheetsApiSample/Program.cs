using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Linq;

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
                

                List<MahjongMember> memberList = new();
                DisplayMember(memberList);

                bool isCreating = true;
                while (isCreating)
                {
                    Console.WriteLine("操作を行う番号を選択してください");
                    Console.WriteLine("1. メンバーの追加");
                    Console.WriteLine("2. 抜き牌データの入力");
                    Console.WriteLine("3. 抜き牌データの作成");
                    Console.Write(">");
                    //Console.WriteLine("1. メンバーの追加");
                    var selection = Console.ReadLine();
                    switch (selection)
                    {
                        case "1":
                            AddMember(memberList);
                            DisplayMember(memberList);
                            break;
                        case "2":
                            break;
                        case "3":
                            CreateGoogleSheets();
                            isCreating = false;
                            break;
                        default:
                            Console.WriteLine("番号が間違っています");
                            Console.WriteLine();
                            break;
                    }
                }
                



                Console.ReadLine();

                //CreateSheet();

                //Console.Write("北(西野)：");
                //var norceCountInNishino =  Console.ReadLine();
                //Console.Write("北(ごっち)：");
                //var norceCountInGotti = Console.ReadLine();
                //Console.Write("北(なる)：");
                //var norceCountInNaru = Console.ReadLine();
                //Console.Write("北(はっしー)：");
                //var norceCountInHasshi = Console.ReadLine();


                
            }
            catch (FileNotFoundException e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private static void DisplayMember(List<MahjongMember> memberList)
        {
            Console.WriteLine("現在のメンバー");
            foreach (var member in memberList.Select((name, index) => new { Value = name, Index = index }))
            {
                Console.WriteLine($"メンバー{member.Index + 1}：{member.Value.Name}");
            }
            Console.WriteLine();
        }

        private static void DisplayExtractedTileInfo(List<MahjongMember> memberList, ExtractedTileInfo extractedTileInfo)
        {
            foreach (var member in memberList)
            {
                Console.WriteLine($"{member.Name}");
            }
            Console.WriteLine("現在の抜き牌情報");
            if (extractedTileInfo == null)
            {
                Console.WriteLine("登録情報なし");
                Console.WriteLine();
                return;
            }
            Console.WriteLine($"タイトル：{extractedTileInfo.Title}");
            Console.WriteLine($"北：{extractedTileInfo.Title}");

        }

        private static void AddMember(List<MahjongMember> memberList)
        {
            MahjongMember member = new();
            Console.Write($"メンバーの名前を入力：");
            member.Name = Console.ReadLine();
            memberList.Add(member);
        }

        //private static void 

        private static void CreateGoogleSheets()
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
                new List<object>{}
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

        private static void CreateSheet()
        {
            Console.Write("タイトル名");
            var titleName = Console.ReadLine();

            List<MahjongMember> memberList = new();
            MahjongMember member = new();
            bool isAddedMember = true;
            int i = 1;
            while (isAddedMember)
            {
                Console.Write($"メンバー{i}の名前を入力：");
                member.Name = Console.ReadLine();
                memberList.Add(member);

                Console.WriteLine($"メンバーの入力を続けますか？(Y/N)");
                var addMember =  Console.ReadLine();
                if (addMember == "Y")
                {
                    i++;
                    continue;
                }
                else if (addMember == "N")
                {
                    break;
                }
                
            }
        }

        class MahjongMember
        {
            public string Name { get; set; }

            public ExtractedTileInfo extractedTileInfo { get; set; }
        }

        class ExtractedTileInfo
        {
            public string Title { get; set; }

            public int NorthCount { get; set; }

            public int SpringCount { get; set; }

            public int SummmerCount { get; set; }

            public int AutumnCount { get; set; }

            public int WinterCount { get; set; }
        }
    }
}