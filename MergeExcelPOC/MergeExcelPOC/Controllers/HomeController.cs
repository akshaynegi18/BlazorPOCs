using ClosedXML.Excel;
using MergeExcelPOC.Models;
using Microsoft.AspNetCore.Mvc;
using System.Data;
using System.Diagnostics;

namespace MergeExcelPOC.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        DataTable table = new DataTable();


        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
            CreateSampleDataTable(900);
        }

        public IActionResult ExportDataTableToExcel(DataTable originalTable)
        {
            try
            {
                // Step 1: Split DataTable
                var chunkSize = 10; // Set your chunk size as needed
                var dataTables = SplitDataTable(originalTable, chunkSize);

                // Step 2: Save each chunk to a temporary file
                var tempFiles = new List<string>();
                foreach (var table in dataTables)
                {
                    var tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
                    SaveDataTableToExcel(table, tempFilePath);
                    tempFiles.Add(tempFilePath);
                }
                // Step 3: Save all chunks to a single Excel file
                var finalFilePath = Path.Combine(Path.GetTempPath(), "FinalExcel.xlsx");
                SaveDataTablesToSingleExcelFile(dataTables, finalFilePath);

                // Return the file as a download
                var bytes = System.IO.File.ReadAllBytes(finalFilePath);
                return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "FinalExcel.xlsx");

                //// Step 3: Merge temporary files into a single Excel file
                //var finalFilePath = Path.Combine(Path.GetTempPath(), "FinalExcel.xlsx");
                //MergeExcelFiles(tempFiles, finalFilePath);

                //// Return the file as a download
                //var bytes = System.IO.File.ReadAllBytes(finalFilePath);
                //return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "FinalExcel.xlsx");
            }
            catch (Exception ex)
            {
                // Handle exceptions
                return StatusCode(500, "Internal server error: " + ex.Message);
            }
        }
        public void MergeExcelFiles(List<string> tempFilePaths, string finalFilePath)
        {
            using (var finalWorkbook = new ClosedXML.Excel.XLWorkbook())
            {
                int sheetIndex = 1;
                foreach (var tempFilePath in tempFilePaths)
                {
                    using (var tempWorkbook = new ClosedXML.Excel.XLWorkbook(tempFilePath))
                    {
                        var tempSheet = tempWorkbook.Worksheet(1);
                        var finalSheet = finalWorkbook.AddWorksheet(tempSheet.Name + sheetIndex);
                        //finalSheet.Cell(1,1).InsertTable(tempSheet.Tables.First());
                        sheetIndex++;
                    }
                }
                finalWorkbook.SaveAs(finalFilePath);
            }
        }
        //public void SaveDataTablesToSingleExcelFile(List<DataTable> dataTables, string filePath)
        //{
        //    using (var workbook = new XLWorkbook())
        //    {
        //        var worksheet = workbook.Worksheets.Add("Sheet1");
        //        int currentRow = 1;

        //        foreach (var table in dataTables)
        //        {
        //            worksheet.Cell(currentRow, 1).InsertTable(table);
        //            currentRow += table.Rows.Count + 1; // Adjust the starting row for the next chunk
        //        }

        //        workbook.SaveAs(filePath);
        //    }
        //}
        public void SaveDataTablesToSingleExcelFile(List<DataTable> dataTables, string filePath)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sheet1");
                int currentRow = 1;
                bool isFirstChunk = true;

                foreach (var table in dataTables)
                {
                    if (isFirstChunk)
                    {
                        worksheet.Cell(currentRow, 1).InsertTable(table); // Include headers
                        currentRow += table.Rows.Count + 1; // Move to the next row after the table
                        isFirstChunk = false;
                    }
                    else
                    {
                        // Copy rows without headers
                        foreach (DataRow row in table.Rows)
                        {
                            for (int col = 0; col < table.Columns.Count; col++)
                            {
                                worksheet.Cell(currentRow, col + 1).Value = row[col].ToString();
                            }
                            currentRow++;
                        }
                    }
                }

                workbook.SaveAs(filePath);
            }
        }


        public static List<DataTable> SplitDataTable(DataTable originalTable, int chunkSize)
        {
            var tables = new List<DataTable>();
            var totalRows = originalTable.Rows.Count;

            for (int i = 0; i < totalRows; i += chunkSize)
            {
                var table = originalTable.Clone();
                for (int j = 0; j < chunkSize && (i + j) < totalRows; j++)
                {
                    table.ImportRow(originalTable.Rows[i + j]);
                }
                tables.Add(table);
            }

            return tables;
        }

        public void SaveDataTableToExcel(DataTable table, string filePath)
        {
            using (var workbook = new ClosedXML.Excel.XLWorkbook())
            {
                workbook.Worksheets.Add(table, "Sheet1");
                workbook.SaveAs(filePath);
            }
        }

        public void CreateSampleDataTable(int rows)
        {
            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Value", typeof(decimal));

            for (int i = 1; i <= rows; i++)
            {
                table.Rows.Add(i, $"Name{i}", i * 1.1m);
            }
            ExportDataTableToExcel(table);

            //return table;
        }


        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}