using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfOpenXMLCURDApplication.Services
{
    // Interface for ExcelService to ensure loose coupling and better testability
    public interface IExcelService
    {
        void CreateExcelFile(string filePath);
        DataTable ReadExcelFile(string filePath);
        void UpdateCell(string filePath, string sheetName, string addressName, string value);
        void DeleteRow(string filePath, string sheetName, uint rowIndex);
    }
}
