using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXml.Excel
{
    public class ExcelCell
    {
        public static String GetCellRawValue(Cell cellObject)
        {
            String CellRawValue = "";
            CellRawValue = cellObject.CellValue.Text;
            return CellRawValue;
        }

        public static String GetCellRawValue(String cellLocation, String sheetname, SpreadsheetDocument spreadsheetObj)
        {
            String CellRawValue = "";
            return CellRawValue;
        }

        public static String GetCellRawValue(String cellLocation, Worksheet sheetObj, SpreadsheetDocument spreadsheetObj)
        {
            String CellRawValue = "";
            return CellRawValue;
        }

        #region Get cell values with strong type support

        public static Boolean GetCellBooleanValue(Cell cellObject)
        {
            Boolean boolCellValue = false;
            String cellRawValue = GetCellRawValue(cellObject);
            if (cellObject.DataType == CellValues.Boolean)
            {
                switch (cellRawValue)
                {
                    case "0":
                        {
                            boolCellValue = false;
                            break;
                        }
                    case "1":
                        {
                            boolCellValue = true;
                            break;
                        }
                    default:
                        break;
                }
            }
            else
            {
                throw new InvalidCastException("Can't cast this cell value to Boolean");
            }
            return boolCellValue;
        }

        public static DateTime GetCellDateValue(Cell cellObject)
        {
            DateTime result = DateTime.Now;
            if (cellObject.DataType.Value == CellValues.Date)
            {
                String cellRaw = ExcelCell.GetCellRawValue(cellObject);
            }
            return result;
        }

        public static String GetCellStringValueByLocation(String cellLocation, String sheetname, SpreadsheetDocument spreadsheetObj)
        {
            String strCellValue = "";
            return strCellValue;
        }

        public static Boolean GetCellBooleanValueByLocation(String cellLocation, String sheetname, SpreadsheetDocument spreadsheetObj)
        {
            Boolean boolCellValue = false;
            return boolCellValue;
        }
        #endregion
    }
}
