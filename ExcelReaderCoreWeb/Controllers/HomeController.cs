using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using ExcelReaderCoreWeb.Models;
using ExcelDataReader;
using System.IO;
using Microsoft.AspNetCore.Http;

namespace ExcelReaderCoreWeb.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult About()
        {
            ViewData["Message"] = "Your application description page.";

            return View();
        }

        public IActionResult Contact()
        {
            ViewData["Message"] = "Your contact page.";

            return View();
        }

        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        [HttpPost]
        public IActionResult ReadFile()
        {
            var files = Request.Form.Files;

            int count = 1, totalRows = 0;
            var errorList = new List<ErrorListModel>();


            for (int i = 0; i < files.Count; i++)
            {
                bool firstRow = false;

                var file = files[i];

                using (var stream = new MemoryStream())
                {
                    file.CopyTo(stream);

                    IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

                    while (excelReader.Read())
                    {
                        if (firstRow)
                        {
                            totalRows++;

                            try
                            {
                                var Column1 = excelReader.GetValue(0);
                                CheckDataType(count, errorList, Column1, typeof(string), 1);

                                var Column2 = excelReader.GetValue(1);
                                CheckDataType(count, errorList, Column2, typeof(string), 2);

                                var Column3 = excelReader.GetValue(2);
                                CheckDataType(count, errorList, Column3, typeof(string), 3);

                                var Column4 = excelReader.GetValue(3);
                                // no type checking
                            }
                            catch (Exception ex)
                            {
                                // Catch general exception...
                            }

                            count++;
                        }
                        else
                        {
                            firstRow = true;
                        }
                    }

                    excelReader.Close();
                }
            }

            if (errorList.Count > 0)
                return Ok(errorList);
            else
                return Ok($"Processed {count} / {totalRows} rows from file.");
        }

        private static void CheckDataType(int count, List<ErrorListModel> errorList, object Value, dynamic DataType, int Column)
        {
            if (Value.GetType() != DataType) errorList.Add(new ErrorListModel { Column = Column, Row = count, Message = $"Value is not of type ({DataType})" });
        }
    }

    public class ErrorListModel
    {
        public int Column { get; set; }
        public int Row { get; set; }
        public string Message { get; set; }
    }
}
