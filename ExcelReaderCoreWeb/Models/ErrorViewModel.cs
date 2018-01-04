using System;

namespace ExcelReaderCoreWeb.Models
{
    public class ErrorListModel
    {
        public int Column { get; set; }
        public int Row { get; set; }
        public string Message { get; set; }
    }
}