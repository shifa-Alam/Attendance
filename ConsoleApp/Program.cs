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
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            //ReadFile();

            //var path1 = @"C:\Users\shifa\Desktop\Attendance.xlsx";
            //var path1 = @"C:\Users\shifa\Desktop\test\attendance.xlsx";
            //ReadFileWithXlReader(path1);


            //FormatDataBl(path1);

            MonitorDirectory(path);

            Console.ReadKey();
        }

        private static string ReadFileWithXlReader( string path)
        {
            try
            {
                //    using (var stream = File.Open(@"C:\Users\shifa\Desktop\attendanceMachineFile.xls", FileMode.Open, FileAccess.Read))
                //    {

                //        using (var reader = ExcelReaderFactory.CreateReader(stream))
                //        {
                //            do
                //            {
                //                while (reader.Read()) //Each ROW
                //                {
                //                    for (int column = 0; column < reader.FieldCount; column++)
                //                    {
                //                        //Console.WriteLine(reader.GetString(column));//Will blow up if the value is decimal etc. 
                //                        Console.WriteLine(reader.GetString(column));//Get Value returns object
                //                    }
                //                }
                //            } while (reader.NextResult()); //Move to NEXT SHEET

                //        }
                //    }
                //


                var attendance = new Attendance();
                //var employeeAttendances = new List<EmployeeAttendance>();
                FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);

                //1. Reading from a binary Excel file ('97-2003 format; *.xls)
                IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                //...
                //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
                //IExcelDataReader excelDReader = ExcelReaderFactory.CreateOpenXmlReader(stream);


                //5. Data Reader methods
                excelReader.Read();
                while (excelReader.Read())
                {
                    var employeeAttendance = new EmployeeAttendance()
                    {
                        EmployeeId = Convert.ToInt64(excelReader.GetString(0)),
                        StartDate = (excelReader.GetString(4) == "C/In") ? Convert.ToDateTime(excelReader.GetString(3)) : null,
                        EndDate = (excelReader.GetString(4) == "C/Out") ? Convert.ToDateTime(excelReader.GetString(3)) : null,

                    };

                    attendance.EmployeeAttendances.Add(employeeAttendance);
                   
                }
                
                //6. Free resources (IExcelDataReader is IDisposable)
                excelReader.Close();
                var json = JsonConvert.SerializeObject(attendance);
                Console.WriteLine(json);
                return json;

            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
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

            //var path1 = @"C:\Users\shifa\Desktop\Attendance.xlsx";
            //var path1 = @"C:\Users\shifa\Desktop\test\attendanceMachineFile.xls";
            //ReadFile();

            //FormatDataBl(path1);
            var jsonData  =ReadFileWithXlReader(e.FullPath);



            _ = CallWebAPIAsync(jsonData);

            Console.WriteLine("File Created :{0}", e.Name);
        }

        private static void ReadFile()
        {
            try
            {
                string con =
                    @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\shifa\Desktop\test\attendanceMachineFile.xls;" +
                    @"Extended Properties='Excel 12.0;HDR=Yes;'";
                using (OleDbConnection connection = new OleDbConnection(con))
                {
                    connection.Open();
                    var dtSchema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                    var Sheet1 = dtSchema.Rows[0].Field<string>("TABLE_NAME").ToString();
                    OleDbCommand command = new OleDbCommand($"select * from [{Sheet1}]", connection);


                    using (OleDbDataReader dr = command.ExecuteReader())
                    {

                        if (dr.Read())
                        {
                            var rows = dr.FieldCount;
                            for (int i = 0; i < rows; i++)
                            {
                                for (int j = 0; j <= i; j++)
                                {
                                    var row1Col0 = dr[i];
                                    Console.WriteLine(row1Col0);
                                }
                            }
                            //while (dr.Read())
                            //{
                            //    var row1Col0 = dr[0];
                            //    Console.WriteLine(row1Col0);
                            //}
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
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
                 

                    var cmd = conn.CreateCommand();
                    cmd.CommandText = String.Format(
                        @"SELECT * FROM [{0}]", Sheet1

                    );

                    using (var rdr = cmd.ExecuteReader())
                    {
                        //LINQ query - when executed will create anonymous objects for each row
                        var query = (from DbDataRecord row in rdr select row).Select(x =>
                            {
                                //dynamic item = new ExpandoObject();
                                Dictionary<string, object> item = new Dictionary<string, object>();
                                for (int i = 0; i < x.FieldCount; i++)
                                {
                                    item.Add(rdr.GetName(i), x[i]);
                                    //item.Add("Abc", x[0]);
                                }


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


        private static async Task CallWebAPIAsync(string jsonData)
        {


            //var opt = new JsonSerializerOptions() { WriteIndented = true };

            //string strJson = JsonSerializer.Serialize<AttendanceJson>(data, opt);


            HttpClient client = new HttpClient();
            client.BaseAddress = new Uri("http://localhost:44317/");
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            //var response = await client.PostAsync("api/HrAttendanceTest/ImportAttendanceAsync", new StringContent(json, Encoding.UTF8, "application/json"));
            var response = await client.PostAsync("api/HrAttendanceTest/SaveAttendanceAsync", new StringContent(jsonData, Encoding.UTF8, "application/json"));

            if (response != null)
            {
                Console.WriteLine(response.ToString());
            }
        }


    }
}

