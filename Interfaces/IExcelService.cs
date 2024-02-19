namespace ExportFile.Interfaces;
public interface IExcelService
{
    Task<byte[]> ConvertExcelToCSV(IFormFile file);
    Task<byte[]> ConvertExcelToJson(IFormFile excelFile);
    Task<byte[]> ConvertExcelToWord(IFormFile excelFile);
    Task<byte[]> ConvertExcelToPdf(IFormFile excelFile);
    Task<byte[]> ConvertExcelToXml(IFormFile excelFile);
}