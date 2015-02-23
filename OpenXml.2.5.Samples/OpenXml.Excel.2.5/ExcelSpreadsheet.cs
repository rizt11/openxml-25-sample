using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXml.Excel
{
    public class ExcelSpreadsheet
    {
        /// <summary>Create simple Excel spreadsheet with basic Workbook.</summary>
        /// <param name="filepath">File path and file name.</param>
        /// <param name="MacroEnabled">Value is true if macro is enabled</param>
        /// <returns></returns>
        public static SpreadsheetDocument CreateExcelDoc(String filepath, bool MacroEnabled)
        {
            SpreadsheetDocument spreadsheet;
            try
            {
                if (MacroEnabled)
                {
                    spreadsheet = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.MacroEnabledWorkbook);
                }
                else
                {
                    spreadsheet = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);
                }
            }
            catch (Exception)
            {
                throw;
            }
            return spreadsheet;
        }

        public static SpreadsheetDocument OpenExcelFile(String filepath)
        {
            SpreadsheetDocument spreadsheet;
            try
            {
                spreadsheet = SpreadsheetDocument.Open(filepath, false);
            }
            catch (Exception)
            {
                throw;
            }
            return spreadsheet;
        }

        public static SpreadsheetDocument OpenExcelFile(String filepath, bool ReadOnly)
        {
            SpreadsheetDocument spreadsheet;
            try
            {
                spreadsheet = SpreadsheetDocument.Open(filepath, !ReadOnly);
            }
            catch (Exception)
            {
                throw;
            }
            return spreadsheet;
        }

        /// <summary>Get current workbook.</summary>
        /// <param name="spreadsheet"></param>
        /// <returns></returns>
        public static Workbook GetCurrentWorkbook(SpreadsheetDocument spreadsheet)
        {
            Workbook currentworkbook = spreadsheet.WorkbookPart.Workbook;
            return currentworkbook;
        }

        /// <summary>Get all sheets in a workbook part.</summary>
        /// <param name="spreadsheet"></param>
        /// <returns></returns>
        public static Sheets GetSheets(SpreadsheetDocument spreadsheet)
        {
            Workbook currentworkbook = GetCurrentWorkbook(spreadsheet);
            var sheets = currentworkbook.Sheets;
            return sheets;
        }

        public static Dictionary<String, String> GetAllNamedRanges(SpreadsheetDocument spreadsheet)
        {
            var RangedNames = new Dictionary<String, String>();
            var wbPart = spreadsheet.WorkbookPart;

            // Retrieve a reference to the defined names collection.
            DefinedNames definedNames = wbPart.Workbook.DefinedNames;

            // If there are defined names, add them to the dictionary.
            if (definedNames != null)
            {
                foreach (DefinedName dn in definedNames)
                    RangedNames.Add(dn.Name.Value, dn.Text);
            }
            return RangedNames;
        }

        #region Manage worksheets and worksheetparts

        public static List<WorksheetPart> GetAllWorksheetParts(SpreadsheetDocument spreadsheet)
        {
            var worksheetParts = new List<WorksheetPart>();
            var sheets = ExcelSpreadsheet.GetSheets(spreadsheet);
            foreach (Sheet itemSheet in sheets)
            {
                String SheetID = itemSheet.Id;
                WorksheetPart wsPart = (WorksheetPart)spreadsheet.WorkbookPart.GetPartById(SheetID);
                worksheetParts.Add(wsPart);
            }
            return worksheetParts;
        }

        public static List<Worksheet> GetAllWorksheets(SpreadsheetDocument spreadsheet)
        {
            var worksheets = new List<Worksheet>();
            var worksheetParts = ExcelSpreadsheet.GetAllWorksheetParts(spreadsheet);
            foreach (WorksheetPart itemWorksheetPart in worksheetParts)
            {
                worksheets.Add(itemWorksheetPart.Worksheet);
            }
            return worksheets;
        }

        #endregion
    }
}
