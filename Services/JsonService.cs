using Newtonsoft.Json.Linq;
using iText.Kernel.Pdf;
using iText.Layout.Element;
using Xceed.Words.NET;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Globalization;
using ExportFile.Interfaces;
using iText.Layout;
using Newtonsoft.Json;

namespace ExportFile.Services;
public class JsonService : IJsonService
{
    public async Task<byte[]> ConvertJsonToPdf(IFormFile jsonFile)
    {
        try
        {
            using (var stream = jsonFile.OpenReadStream())
            {
                using (var reader = new StreamReader(stream))
                {
                    var jsonData = await reader.ReadToEndAsync();
                    var jsonArray = JArray.Parse(jsonData);

                    using (var memoryStream = new MemoryStream())
                    {
                        var writer = new PdfWriter(memoryStream);
                        var pdf = new PdfDocument(writer);
                        var document = new Document(pdf);
                        foreach (var token in jsonArray)
                        {
                            if (token is JObject jsonObject)
                            {
                                foreach (var property in jsonObject.Properties())
                                {
                                    document.Add(new Paragraph($"{property.Name}: {property.Value}"));
                                }
                                document.Add(new Paragraph(""));
                            }
                        }
                        document.Close();
                        return memoryStream.ToArray();
                    }
                }
            }
        }
        catch (Exception ex)
        {
            throw new Exception("Error converting JSON to PDF: " + ex.Message);
        }
    }

    public async Task<byte[]> ConvertJsonToWord(IFormFile jsonFile)
    {
        try
        {
            using (var streamReader = new StreamReader(jsonFile.OpenReadStream()))
            {
                var jsonData = await streamReader.ReadToEndAsync();
                JArray jsonArray = JArray.Parse(jsonData);
                using (DocX document = DocX.Create("output.docx"))
                {
                    foreach (JObject jsonObject in jsonArray)
                    {
                        foreach (var property in jsonObject)
                        {
                            document.InsertParagraph($"{property.Key}: {property.Value}");
                        }
                        document.InsertParagraph("");
                    }
                    document.Save();
                }
                return File.ReadAllBytes("output.docx");
            }
        }
        catch (Exception ex)
        {
            throw new Exception("Error converting JSON to Word: " + ex.Message);
        }
    }

    public async Task<byte[]> ConvertJsonToExcel(IFormFile jsonFile)
    {
        try
        {
            using (var streamReader = new StreamReader(jsonFile.OpenReadStream()))
            {
                var jsonData = await streamReader.ReadToEndAsync();
                List<Dictionary<string, object>> dataList = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(jsonData);
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("Sheet1");
                IRow headerRow = sheet.CreateRow(0);
                int cellIndex = 0;
                foreach (var key in dataList[0].Keys)
                {
                    headerRow.CreateCell(cellIndex).SetCellValue(key);
                    cellIndex++;
                }
                int rowIndex = 1;
                foreach (var data in dataList)
                {
                    IRow dataRow = sheet.CreateRow(rowIndex);
                    cellIndex = 0;
                    foreach (var value in data.Values)
                    {
                        dataRow.CreateCell(cellIndex).SetCellValue(value.ToString());
                        cellIndex++;
                    }
                    rowIndex++;
                }
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    workbook.Write(memoryStream);
                    return memoryStream.ToArray();
                }
            }
        }
        catch (Exception ex)
        {
            throw new Exception("Error converting JSON to Excel: " + ex.Message);
        }
    }

    public async Task<byte[]> ConvertJsonToCsv(IFormFile jsonFile)
    {
        try
        {
            using (var streamReader = new StreamReader(jsonFile.OpenReadStream()))
            {
                string jsonContent = await streamReader.ReadToEndAsync();
                JArray jsonArray = JArray.Parse(jsonContent);

                if (jsonArray.Count == 0)
                {
                    throw new Exception("JSON array is empty.");
                }

                string csvFilePath = Path.Combine(Path.GetTempPath(), "output.csv");
                using (var writer = new StreamWriter(csvFilePath))
                {
                    using (var csv = new CsvHelper.CsvWriter(writer, CultureInfo.InvariantCulture))
                    {
                        foreach (JObject item in jsonArray)
                        {
                            foreach (JProperty property in item.Properties())
                            {
                                csv.WriteField(property.Value.ToString());
                            }
                            csv.NextRecord();
                        }
                    }
                }

                return File.ReadAllBytes(csvFilePath);
            }
        }
        catch (Exception ex)
        {
            throw new Exception("Error converting JSON to CSV: " + ex.Message);
        }
    }
}