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

namespace test_conn
{
   public class Doc_conn
    {
      //  https://docs.google.com/spreadsheets/d/1w4FXbrqr0okwwRSEqoF7SdJKxDIqqIltXSQgpklqylM/edit?usp=sharing
        public static String spreadsheetId = "1w4FXbrqr0okwwRSEqoF7SdJKxDIqqIltXSQgpklqylM";
        public static SheetsService connection;

        public Doc_conn()
        {
            if(connection is null) connection = Get_Connection();
        }

        public void Update(String range, ValueRange Content)
        {
         
            Google.Apis.Sheets.v4.Data.ValueRange requestBody = new Google.Apis.Sheets.v4.Data.ValueRange();
            var values  = new List<IList<object>>();
            values.Add(new List<object> { "Jacek", "Pawel" });
            requestBody.Values = values;
            requestBody.Range = "A2:A3";

            var request = connection.Spreadsheets.Values.Append(requestBody, spreadsheetId, range);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;

            var response = request.Execute();
            Console.WriteLine(JsonConvert.SerializeObject(response));
        }
        public  ValueRange Get(String range)
        {
            SpreadsheetsResource.ValuesResource.GetRequest request = connection.Spreadsheets.Values.Get(spreadsheetId, range);
            return  request.Execute();
        }


        private static SheetsService Get_Connection()
        {
            // If modifying these scopes, delete your previously saved credentials
            // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
           
            string[] Scopes = { "https://www.googleapis.com/auth/spreadsheets" };
            string ApplicationName = "Google Sheets API .NET Quickstart";
            GoogleCredential credential;

            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {

                /*string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;*/

                credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
           
            }

            // Create Google Sheets API service.
            return new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

        }
              
             
            

        
    }
}