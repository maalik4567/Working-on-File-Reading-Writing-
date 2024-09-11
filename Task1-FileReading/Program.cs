using System;
using System.Collections.Generic;
using System.Data.SqlClient; // For SQL connection
using System.IO;
using Newtonsoft.Json;
using OfficeOpenXml;

namespace FileWorking
{
    class Program
    {
        // TASK-01 Method to write data to Excel
        static void WriteToExcel(List<ColorInfo> colorData, string excelFilePath)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Colors");

                // Add headers
                worksheet.Cells[1, 1].Value = "Color";
                worksheet.Cells[1, 2].Value = "Value";

                // Add data to Excel
                int rowIndex = 2; // Start at row 2 because row 1 is for headers
                foreach (var colorInfo in colorData)
                {
                    worksheet.Cells[rowIndex, 1].Value = colorInfo.Color;
                    worksheet.Cells[rowIndex, 2].Value = colorInfo.Value;
                    rowIndex++;
                }

                // Save the Excel file to the specified path
                package.SaveAs(new FileInfo(excelFilePath));
            }

            Console.WriteLine("\nData has been written to Excel successfully! Check your Desktop for 'colors_output.xlsx'.");
        }

        //TASK-02 Method to read data from Excel and print to console
        static void ReadFromExcel(string excelFilePath)
        {
            if (!File.Exists(excelFilePath))
            {
                Console.WriteLine("Excel file not found at the specified path: " + excelFilePath);
                return;
            }

            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0]; // Assuming it's the first worksheet
                int rowCount = worksheet.Dimension.Rows;        // Get total number of rows
                int colCount = worksheet.Dimension.Columns;     // Get total number of columns

                Console.WriteLine("\nReading from Excel:");
                for (int row = 1; row <= rowCount; row++)
                {
                    for (int col = 1; col <= colCount; col++)
                    {
                        Console.Write(worksheet.Cells[row, col].Text + "\t"); // Print the cell text
                    }
                    Console.WriteLine();
                }
            }

            Console.WriteLine("\nData has been read from Excel and printed to console.");
        }

        //TASK-03  Method to insert data into the database
        static void InsertDataIntoDatabase(string excelFilePath)
        {
            // Your SQL connection string (update with your actual database details)
            string connectionString = "Data Source=DESKTOP-P7354T7\\SQLEXPRESS;Initial Catalog=DataOpt;Integrated Security=True";

            // Ensure the Excel file exists
            if (!File.Exists(excelFilePath))
            {
                Console.WriteLine("Excel file not found at the specified path: " + excelFilePath);
                return;
            }

            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0]; // Assuming it's the first worksheet
                int rowCount = worksheet.Dimension.Rows;        // Get total number of rows

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    for (int row = 2; row <= rowCount; row++) // Start from row 2 to skip headers
                    {
                        string color = worksheet.Cells[row, 1].Text;
                        string value = worksheet.Cells[row, 2].Text;

                        // SQL query to insert data into ExcelData table
                        string query = "INSERT INTO ExcelData (color, value) VALUES (@color, @value)";

                        using (SqlCommand command = new SqlCommand(query, connection))
                        {
                            command.Parameters.AddWithValue("@color", color);
                            command.Parameters.AddWithValue("@value", value);

                            // Execute the query
                            command.ExecuteNonQuery();
                        }
                    }

                    Console.WriteLine("\nData has been successfully inserted into the database.");
                }
            }
        }

        // Method to insert data from the database into an Excel file
        static void InsertDataIntoExcelFromDB(string excelFilePath)
        {
            string connectionString = "Data Source=DESKTOP-P7354T7\\SQLEXPRESS;Initial Catalog=DataOpt;Integrated Security=True";
            string query = "SELECT color, value FROM ExcelData"; // Adjust table and column names if needed

            // Create an Excel package and add a worksheet
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("DataFromDB");

                // Add headers
                worksheet.Cells[1, 1].Value = "Color";
                worksheet.Cells[1, 2].Value = "Value";

                int rowIndex = 2;

                using (var connection = new SqlConnection(connectionString))
                using (var command = new SqlCommand(query, connection))
                {
                    connection.Open();

                    using (var reader = command.ExecuteReader())
                    {
                        // Read data from the database and write to Excel
                        while (reader.Read())
                        {
                            worksheet.Cells[rowIndex, 1].Value = reader["color"].ToString();
                            worksheet.Cells[rowIndex, 2].Value = reader["value"].ToString();
                            rowIndex++;
                        }
                    }
                }

                // Save the Excel file to the specified path
                package.SaveAs(new FileInfo(excelFilePath));
            }

            Console.WriteLine("\nData has been successfully written to the Excel file.");
        }

        // Method to convert SQL data to a JSON file
        static void ConvertDbToJson(string jsonFilePath)
        {
            // Define the connection string and SQL query inside the method
            string connectionString = "Data Source=DESKTOP-P7354T7\\SQLEXPRESS;Initial Catalog=DataOpt;Integrated Security=True";
            string query = "SELECT color, value FROM ExcelData"; // Adjust the SQL query as needed

            var colorData = new List<ColorInfo>();

            // Connect to the database and retrieve data
            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();

                using (var command = new SqlCommand(query, connection))
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        // Create a ColorInfo object for each row
                        var colorInfo = new ColorInfo
                        {
                            Color = reader["color"].ToString(),
                            Value = reader["value"].ToString()
                        };
                        colorData.Add(colorInfo);
                    }
                }
            }

            // Serialize the list to JSON
            string jsonData = JsonConvert.SerializeObject(colorData, Formatting.Indented);

            // Write JSON to file
            File.WriteAllText(jsonFilePath, jsonData);

            Console.WriteLine("Data has been successfully written to JSON file.");
        }

        // Class to map the JSON data
        public class ColorInfo
        {
            public string Color { get; set; }
            public string Value { get; set; }
        }

        static void Main(string[] args)
        {
            // Set the license context for EPPlus (non-commercial)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Hardcoded path for the JSON file on Desktop
            string jsonFilePath = @"C:/Users/User/Desktop/data.json";

            // Check if the file exists
            if (!File.Exists(jsonFilePath))
            {
                Console.WriteLine("JSON file not found at the specified path: " + jsonFilePath);
                return;
            }

            // Read the JSON file
            string jsonData = File.ReadAllText(jsonFilePath);

            // Deserialize JSON to a list of objects (array of objects in JSON)
            var colorData = JsonConvert.DeserializeObject<List<ColorInfo>>(jsonData);

            // Hardcoded path for the Excel file to be created on Desktop
            string excelFilePath = @"C:/Users/User/Desktop/colors_output.xlsx";

            // Write data to Excel
            WriteToExcel(colorData, excelFilePath);

            // Now read from the Excel file and print to console
            ReadFromExcel(excelFilePath);

            // Insert the data from Excel into the database
            InsertDataIntoDatabase(excelFilePath);

            // DB TO EXCEL
            string excelfilePath2 = @"C:/Users/User/Desktop/DataFromDB.xlsx";
            InsertDataIntoExcelFromDB(excelfilePath2);

            // DB TO JSON 
            string jsonFilePath2 = @"C:/Users/User/Desktop/dataset-colors.json"; // Path to save the JSON file

            ConvertDbToJson(jsonFilePath2);

        }
    }
}
