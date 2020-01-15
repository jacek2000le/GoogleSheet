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
    class Program
    {

        public static Doc_conn conn;
        static void Main(string[] args)
        {
            conn = new Doc_conn();
             conn.Update("A2:A3", null);
            
            /*IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                 Console.WriteLine("{0}",row[0]);
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }
            Console.Read();*/
        }
    }
}