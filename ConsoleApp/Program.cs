using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using ExcelDataReader;
using Newtonsoft.Json;
using JsonSerializer = System.Text.Json.JsonSerializer;

namespace ConsoleApp
{
    class Program
    {
        static string path = @"C:\Users\shifa\Desktop\test";

        static void Main(string[] args)
        {
            MonitorDirectory(path);
            Console.ReadKey();
        }
        private static void MonitorDirectory(string path)
        {
            FileSystemWatcher fileSystemWatcher = new FileSystemWatcher();
            fileSystemWatcher.Path = path;
            fileSystemWatcher.Created += FileSystemWatcher_Created;
            //fileSystemWatcher.Renamed += FileSystemWatcher_Renamed;
            //fileSystemWatcher.Deleted += FileSystemWatcher_Deleted;
            fileSystemWatcher.EnableRaisingEvents = true;


        }



        private static void FileSystemWatcher_Created(object sender, FileSystemEventArgs e)
        {

            var path1 = @"C:\Users\shifa\Desktop\Attendance.xlsx";
            //var path1 = @"C:\Users\shifa\Desktop\test\attendanceMachineFile.xls";

            FormatDataBl(path1);



            //_ = CallWebAPIAsync(excel);

            Console.WriteLine("File Created :{0}", e.Name);
        }

        private static object FormatDataBl(string path1)
        {
            try
            {

                //var pathToExcel = path1;
                //var sheetName = "Sheet1";

                //This connection string works if you have Office 2007+ installed and your 
                //data is saved in a .xlsx file
                var connectionString = String.Format(@"
                    Provider=Microsoft.ACE.OLEDB.12.0;
                    Data Source={0};
                    Extended Properties=""Excel 12.0 Xml;HDR=YES""", path1);

                //Creating and opening a data connection to the Excel sheet 
                using (var conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    var dtSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                    var Sheet1 = dtSchema.Rows[0].Field<string>("TABLE_NAME").ToString();
                    Sheet1 = Sheet1.Remove(Sheet1.Length - 1);

                    var cmd = conn.CreateCommand();
                    cmd.CommandText = String.Format(
                        @"SELECT * FROM [{0}$]", Sheet1

                    );


                    using (var rdr = cmd.ExecuteReader())
                    {
                        //LINQ query - when executed will create anonymous objects for each row
                        var query = (from DbDataRecord row in rdr select row).Select(x =>
                            {


                                //dynamic item = new ExpandoObject();
                                Dictionary<string, object> item = new Dictionary<string, object>();
                                item.Add("Abc", x[0]);
                                item.Add(rdr.GetName(1), x[1]);
                                item.Add(rdr.GetName(2), x[2]);
                                item.Add(rdr.GetName(3), x[3]);
                                item.Add(rdr.GetName(4), x[4]);
                                item.Add(rdr.GetName(5), x[5]);
                                item.Add(rdr.GetName(6), x[6]);
                                item.Add(rdr.GetName(7), x[7]);
                                return item;

                            });

                        //Generates JSON from the LINQ query
                        var json = JsonConvert.SerializeObject(query);
                        return json;
                    }
                }

            }
            catch (Exception e)
            {

                throw;
            }
        }


        private static void FileSystemWatcher_Renamed(object sender, FileSystemEventArgs e)
        {
            Console.WriteLine("File Renamed :{0}", e.Name);
        }

        private static void FileSystemWatcher_Deleted(object sender, FileSystemEventArgs e)
        {
            Console.WriteLine("File Deleted :{0}", e.Name);
        }


        private static async Task CallWebAPIAsync(AttendanceJson data)
        {


            var opt = new JsonSerializerOptions() { WriteIndented = true };

            string strJson = JsonSerializer.Serialize<AttendanceJson>(data, opt);


            HttpClient client = new HttpClient();
            client.BaseAddress = new Uri("http://localhost:44317/");
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            var response = await client.PostAsync("api/HrAttendanceTest/ImportAttendanceAsync", new StringContent(strJson, Encoding.UTF8, "application/json"));

            if (response != null)
            {
                Console.WriteLine(response.ToString());
            }
        }


    }
}

