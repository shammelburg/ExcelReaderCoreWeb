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

        public IActionResult Error()
        {
            return View();
        }

        public IActionResult Success()
        {
            return View();
        }

        [HttpPost]
        public IActionResult ReadFile()
        {
            var files = Request.Form.Files;

            int row = 1, totalRows = 0;
            var errorList = new List<ErrorListModel>();
            var list = new List<Object>();

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
                        int columnNumber = 0;

                        if (firstRow)
                        {
                            totalRows++;
                            
                            try
                            {
                                columnNumber = 1;
                                var Column1 = excelReader.GetString(0).ToString();

                                columnNumber = 2;
                                var Column2 = excelReader.GetString(1).ToString();

                                columnNumber = 3;
                                var Column3 = excelReader.GetDouble(2);

                                columnNumber = 4;
                                var Column4 = excelReader.GetBoolean(3);


                                // Run db command here...
                                // Not recommended for a huge collection
                                // Or
                                // Create another collection and submit that to the db
                                // Using a UDT or XML

                                list.Add(new { c1 = Column1, c2 = Column2, c3 = Column3, c4 = Column4 });
                            }
                            catch (NullReferenceException ex)
                            {
                                errorList.Add(new ErrorListModel { Column = columnNumber, Row = row, Message = $"No Data Found", Exception = ex.Message });
                            }
                            catch (Exception ex)
                            {
                                errorList.Add(new ErrorListModel { Column = columnNumber, Row = row, Message = $"Incorrect Data Type", Exception = ex.Message });
                            }
                        }
                        else
                        {
                            firstRow = true;
                        }

                        row++;
                    }

                    excelReader.Close();
                }
            }

            if (errorList.Count > 0)
                return View("Error", errorList);

            return View("Success", list);
        }
    }
}
