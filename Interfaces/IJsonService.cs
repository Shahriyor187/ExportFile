namespace ExportFile.Interfaces;
public interface IJsonService
{
    Task<byte[]> ConvertJsonToPdf(IFormFile jsonFile);
    Task<byte[]> ConvertJsonToWord(IFormFile jsonFile);
    Task<byte[]> ConvertJsonToExcel(IFormFile jsonFile);
    Task<byte[]> ConvertJsonToCsv(IFormFile jsonFile);
}