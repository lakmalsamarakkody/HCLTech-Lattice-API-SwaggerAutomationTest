using System;
using System.Collections;
using System.IO;
using xl = NPOI.XSSF.UserModel;
using xs = NPOI.SS.UserModel;

namespace SwaggerWebAPI.Libs
{
    class ExcelAPI
    {
        xs.IWorkbook iWorkBook;
        xs.ISheet iSheet;
        xs.IRow iRow;
        xs.ICell iCell;
        Hashtable sheets;

        public String xlfilePath;

        public ExcelAPI(string xlpath)
        {
            this.xlfilePath = xlpath;
            OpenExcel();
        }

        public void OpenExcel()
        {
            iWorkBook = new xl.XSSFWorkbook(File.Open(xlfilePath, FileMode.Open));
            sheets = new Hashtable();
            int noOfSheets = iWorkBook.NumberOfSheets;
            for (int index = 0; index < noOfSheets; index++)
            {
                sheets[index] = iWorkBook.GetSheetName(index);
            }
        }
        public bool isSheetExists(int sheetIndex)
        {
            bool isSheetExists = false;
            if (sheets.ContainsKey(sheetIndex))
            {
                isSheetExists = true;
            }
            return isSheetExists;
        }
        public bool isSheetExists(string sheetName)
        {
            bool isSheetExists = false;
            if (sheets.ContainsValue(sheetName))
            {
                isSheetExists = true;
            }
            return isSheetExists;
        }
        public int GetSheetIndex(string sheetName)
        {
            int SheetIndex = 0;
            if (isSheetExists(sheetName))
            {
                foreach (DictionaryEntry sheet in sheets)
                {
                    if (sheet.Value.Equals(sheetName))
                    {
                        SheetIndex = (int)sheet.Key;
                    }
                }
            }
            else
            {
                throw new Exception();
            }
            return SheetIndex;
        }

        public string GetSheetName(int SheetIndex)
        {
            string sheetName = string.Empty;
            foreach (DictionaryEntry sheet in sheets)
            {
                if (sheet.Key.Equals(SheetIndex))
                {
                    sheetName = sheet.Value.ToString();
                }
            }
            return sheetName;
        }

        public int GetRowCount(string sheetName)
        {
            int rowCount = 0;
            rowCount = GetLastRowIndex(sheetName) + 1;
            return rowCount;
        }

        public int GetLastRowIndex(string sheetName)
        {
            int rowCount = 0;
            if (isSheetExists(sheetName))
            {
                iSheet = iWorkBook.GetSheetAt(GetSheetIndex(sheetName));
                rowCount = iSheet.LastRowNum;
            }
            return rowCount;
        }

        public int GetColumnCountByHeader(string sheetName)
        {
            int colCount = 0;
            colCount = GetColumnCountByRow(sheetName, 0);
            return colCount;
        }

        public int GetColumnCountByRow(string sheetName, int row)
        {
            int colCount = 0;
            if (isSheetExists(sheetName))
            {
                iSheet = iWorkBook.GetSheetAt(GetSheetIndex(sheetName));
                colCount = iSheet.GetRow(row).LastCellNum;
            }
            return colCount;
        }

        // GET CELL DATA
        public string GetCellData(int row, int col)
        {
            return GetCellData(0, row, col);
        }
        public string GetCellData(string sheetName, int row, int col)
        {
            return GetCellData(GetSheetIndex(sheetName), row, col);
        }
        public string GetCellData(int sheetIndex, int row, int col)
        {
            string value = string.Empty;
            if (isSheetExists(sheetIndex))
            {
                iSheet = iWorkBook.GetSheetAt(sheetIndex);
                iRow = iSheet.GetRow(row);
                iCell = iRow.GetCell(col);
                string dataType = iCell.CellType.ToString();

                switch (dataType)
                {
                    case "String":
                        value = iCell.StringCellValue.ToString();
                        break;
                    case "Numeric":
                        value = iCell.NumericCellValue.ToString();
                        break;
                    default:
                        break;
                }
            }
            return value;
        }

        // SET CELL DATA
        public bool SetCellData(string cellValue)
        {
            return SetCellData(0, 0, 0, cellValue);
        }
        public bool SetCellData(int row, int col, string cellValue)
        {
            return SetCellData(0, row, col, cellValue);
        }
        public bool SetCellData(string sheetName, int row, int col, string cellValue)
        {
            return SetCellData(GetSheetIndex(sheetName), row, col, cellValue);
        }
        public bool SetCellData(int sheetIndex, int row, int col, string cellValue)
        {
            if (isSheetExists(sheetIndex))
            {
                iSheet = iWorkBook.GetSheetAt(sheetIndex);
                iRow = iSheet.GetRow(row);
                if (iRow == null)
                {
                    iRow = iSheet.CreateRow(row);
                }
                iCell = iRow.CreateCell(col);
                iCell.SetCellValue(cellValue);

                FileStream OutFileStream = new FileStream(xlfilePath, FileMode.Create, FileAccess.Write);
                //fileStream.Position = 0;
                iWorkBook.Write(OutFileStream, false);
                OutFileStream.Close();
            }
            return true;
        }
        public void CloseExcel()
        {
            iWorkBook.Close();
            iWorkBook = null;
            iSheet = null;
            iRow = null;
            iCell = null;
        }
    }
}
