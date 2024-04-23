using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelFormularValueReaderSampleApp
{
    public static class ExcelReaderService
    {
        public static void ReadExcel(string filePath, int sheetNo)
        {
            try
            {
                IWorkbook wb = new XSSFWorkbook(filePath);
                ISheet ws = wb.GetSheetAt(sheetNo);
                IRow row = ws.GetRow(0);
                ICell cell = row.GetCell(2);
                var cellValue = cell.NumericCellValue;
                Console.WriteLine(cellValue.ToString());
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}