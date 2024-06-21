using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System.Data;

namespace Converter
{
    public class JsonToExcelConverter
    {

        // Method to convert JSON to DataTable
        public static DataTable JsonToDataTable(string json)
        {
            var table = new DataTable();

            JArray jsonArray;
            try
            {
                try
                {
                    jsonArray = JArray.Parse(json);
                }
                catch (JsonReaderException)
                {
                    var jsonObject = JObject.Parse(json);
                    jsonArray = jsonObject.Descendants().Where(d => d is JArray).FirstOrDefault() as JArray;
                }


                if (jsonArray == null)
                {
                    throw new Exception("The provided JSON does not contain an array.");
                }
                return ConvertJsonArrayToTable(table, jsonArray);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error encountered while reading json file. {ex.Message}");
            }
        }

        private static DataTable ConvertJsonArrayToTable(DataTable table, JArray jsonArray)
        {
            foreach (JObject obj in jsonArray.Children<JObject>())
            {
                foreach (JProperty prop in obj.Properties())
                {
                    if (!table.Columns.Contains(prop.Name))
                    {
                        table.Columns.Add(prop.Name, typeof(string));
                    }
                }
            }

            foreach (JObject obj in jsonArray.Children<JObject>())
            {
                var row = table.NewRow();
                foreach (JProperty prop in obj.Properties())
                {
                    row[prop.Name] = prop.Value.ToString();
                }
                table.Rows.Add(row);
            }

            return table;
        }

        // Method to convert DataTable to Excel file
        public static void DataTableToExcel(DataTable dataTable, string excelFilePath)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = dataTable.Columns[i].ColumnName;
                    }

                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataTable.Columns.Count; j++)
                        {
                            worksheet.Cells[i + 2, j + 1].Value = dataTable.Rows[i][j];
                        }
                    }

                    package.SaveAs(new System.IO.FileInfo(excelFilePath));
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error encountered while generating excel file. {ex.Message}");
            }
        }

        // Method to convert JSON directly to Excel file
        public static void JsonToExcel()
        {
            try
            {
                Console.WriteLine("Please note that the excel file will be generated in the same directory as the json file.");
                Console.WriteLine("Sample absolute file path on windows - C:\\Users\\Ozioma Ene\\Projects\\UserLogs.json");
                Console.WriteLine("Kindly insert the absolute file path of the json file:");
                string absoluteFilePath = Console.ReadLine();
                if (string.IsNullOrEmpty(absoluteFilePath))
                {
                    Console.WriteLine(new ArgumentNullException(nameof(string.Empty)));
                    Console.WriteLine("File path cannot be null or empty");
                    return;
                }
                if (!File.Exists(absoluteFilePath))
                {
                    Console.WriteLine(new FileNotFoundException(nameof(string.Empty)));
                    Console.WriteLine("File does not exist.");
                    return;
                }
                DataTable dataTable = JsonToDataTable(File.ReadAllText(absoluteFilePath));
                DataTableToExcel(dataTable, GetOutputFilePath(absoluteFilePath));
                Console.WriteLine($"File has been successfully generated on this path {GetOutputFilePath(absoluteFilePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error encountered while converting file. {ex.Message}");
            }
            finally
            {
                Console.WriteLine("You may rerun the application...");
            }
        }

        public static string GetOutputFilePath(string absoluteFilePath)
        {
            string directoryPath = Path.GetDirectoryName(absoluteFilePath);
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(absoluteFilePath);
            return Path.Combine(directoryPath, fileNameWithoutExtension + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xlsx");
        }

    }
}
