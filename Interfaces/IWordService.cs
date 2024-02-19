namespace ExportFile.Interfaces;
public interface IWordService
{
    Task<byte[]> ConvertWordToJson(IFormFile wordfile);
    Task<byte[]> ConvertWordToExcel(IFormFile wordFile);
}