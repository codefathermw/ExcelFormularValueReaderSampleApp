namespace ExcelFormularValueReaderSampleApp
{
    public class Program
    {
        static void Main(string[] args)
        {
            var filePath = "C:\\Users\\Public\\Files\\sample.xlsx"; // give it your path here
            var sheetNo = 0;
            ExcelReaderService.ReadExcel(filePath, sheetNo);
        }
    }
}