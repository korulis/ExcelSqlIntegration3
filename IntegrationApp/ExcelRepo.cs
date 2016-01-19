using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace IntegrationApp
{
    public class ExcelRepo
    {
        private FileStream _excelFile; 

        public ExcelRepo(FileStream excelFile)
        {
            _excelFile = excelFile;
            //var spreadsheetDocument = SpreadsheetDocument.Open("", false);
            //spreadsheetDocument.Dispose();
        }

        public object GetDataFromExcel(string worksheetName)
        {
            var data = new List<List<string>>();
            using (var spreadsheetDocument = SpreadsheetDocument.Open(_excelFile, false))
            {
                var workbookPart = spreadsheetDocument.WorkbookPart;
                var theSheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == worksheetName);
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(theSheet?.Id);
                var theCells = worksheetPart.Worksheet.Descendants<Cell>().ToList();

                var rawRange = new List<string>();
                var tempList = new List<string>();
                foreach (var c in theCells)
                {
                    const string pattern = @"(([A-Z]+)(1\b))";
                    if (Regex.Matches(c.CellReference, pattern, RegexOptions.IgnoreCase).Count != 0)
                    {
                        continue;
                    }


                    data.Add(new List<string> {c.InnerText});
                    tempList.Add(c.InnerText);
                    rawRange.Add(c.CellReference);
                }
            }
            return data;

        }
    }
}
