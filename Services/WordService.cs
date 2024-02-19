using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ExportFile.Interfaces;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.Text;

namespace ExportFile.Services;
public class WordService : IWordService
{
    public async Task<byte[]> ConvertWordToExcel(IFormFile wordFile)
    {
        try
        {
            if (wordFile == null || wordFile.Length == 0)
                throw new ArgumentException("No file was uploaded");

            using (var stream = wordFile.OpenReadStream())
            {
                using (var wordDocument = WordprocessingDocument.Open(stream, false))
                {
                    var package = new ExcelPackage();
                    var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                    int row = 1;
                    foreach (var paragraph in wordDocument.MainDocumentPart.Document.Body.Elements<Paragraph>())
                    {
                        var text = new StringBuilder();
                        foreach (var run in paragraph.Elements<Run>())
                        {
                            text.Append(run.InnerText);
                        }
                        worksheet.Cells[row++, 1].Value = text.ToString();
                    }

                    byte[] excelBytes = package.GetAsByteArray();
                    return excelBytes;
                }
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Error converting Word to Excel: {ex.Message}");
        }
    }
    public async Task<byte[]> ConvertWordToJson(IFormFile wordfile)
    {
        try
        {
            if (wordfile == null || wordfile.Length == 0)
                throw new ArgumentException("No file was uploaded");

            using (var stream = wordfile.OpenReadStream())
            {
                using (var wordDocument = WordprocessingDocument.Open(stream, false))
                {
                    StringBuilder sb = new StringBuilder();
                    foreach (var paragraph in wordDocument.MainDocumentPart.Document.Body.Elements<Paragraph>())
                    {
                        sb.AppendLine(paragraph.InnerText);
                    }
                    var jsonContent = JsonConvert.SerializeObject(sb.ToString());
                    return Encoding.UTF8.GetBytes(jsonContent);
                }
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Error converting Word to JSON: {ex.Message}");
        }
    }
}