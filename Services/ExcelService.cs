using ClosedXML.Excel;
using ExportFile.Interfaces;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.Data;
using System.Text;

namespace ExportFile.Services;
public class ExcelService : IExcelService
{
    public async Task<byte[]> ConvertExcelToCSV(IFormFile file)
    {
        try
        {
            using (var stream = new MemoryStream())
            {
                await file.CopyToAsync(stream);
                using (ExcelPackage package = new ExcelPackage(stream))
                {
                    if (package.Workbook.Worksheets.Count == 0)
                    {
                        throw new InvalidDataException("The Excel file does not contain any worksheets");
                    }
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    StringBuilder csvString = new StringBuilder();
                    for (int row = 1; row <= worksheet.Dimension.Rows; row++)
                    {
                        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        {
                            if (col > 1)
                                csvString.Append(",");
                            var cellValue = worksheet.Cells[row, col].Value;
                            if (cellValue != null)
                            {
                                string cellText = cellValue.ToString().Replace("\"", "\"\"");
                                if (cellText.Contains(",") || cellText.Contains("\"") ||
                                    cellText.Contains("\r") || cellText.Contains("\n"))
                                {
                                    cellText = $"\"{cellText}\"";
                                }
                                csvString.Append(cellText);
                            }
                        }
                        csvString.AppendLine();
                    }
                    return Encoding.UTF8.GetBytes(csvString.ToString());
                }
            }
        }
        catch (Exception ex)
        {
            throw new Exception("Error converting Excel to CSV: " + ex.Message);
        }
    }

    public async Task<byte[]> ConvertExcelToJson(IFormFile excelFile)
    {
        try
        {
            if (excelFile == null || excelFile.Length == 0)
                throw new ArgumentException("Excel file is empty or null");

            var dataTable = ReadExcelToDataTable(excelFile);
            string jsonOutput = JsonConvert.SerializeObject(dataTable, Newtonsoft.Json.Formatting.Indented);
            return Encoding.UTF8.GetBytes(jsonOutput);
        }
        catch (Exception ex)
        {
            throw new Exception("Error converting Excel to JSON: " + ex.Message);
        }
    }

    public async Task<byte[]> ConvertExcelToPdf(IFormFile excelFile)
    {
        try
        {
            if (excelFile == null || excelFile.Length == 0)
                throw new ArgumentException("Excel file is empty or null");

            var pdfFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".pdf");

            using (var stream = new FileStream(pdfFilePath, FileMode.Create))
            {
                var dataTable = ReadExcelToDataTable(excelFile);

                using (var document = new Document())
                {
                    PdfWriter.GetInstance(document, stream);
                    document.Open();

                    var table = new PdfPTable(dataTable.Columns.Count);
                    foreach (DataColumn column in dataTable.Columns)
                        table.AddCell(new Phrase(column.ColumnName));

                    foreach (DataRow row in dataTable.Rows)
                        foreach (var cell in row.ItemArray)
                            table.AddCell(new Phrase(cell.ToString()));

                    document.Add(table);
                }
            }
            return await File.ReadAllBytesAsync(pdfFilePath);
        }
        catch (Exception ex)
        {
            throw new Exception("Error converting Excel to PDF: " + ex.Message);
        }
    }

    public async Task<byte[]> ConvertExcelToWord(IFormFile excelFile)
    {
        try
        {
            if (excelFile == null || excelFile.Length == 0)
                throw new ArgumentException("Excel file is empty or null");

            var wordFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");

            using (var workbook = new XLWorkbook(excelFile.OpenReadStream()))
            {
                workbook.SaveAs(wordFilePath);
            }

            return await File.ReadAllBytesAsync(wordFilePath);
        }
        catch (Exception ex)
        {
            throw new Exception("Error converting Excel to Word: " + ex.Message);
        }
    }

    public async Task<byte[]> ConvertExcelToXml(IFormFile excelFile)
    {
        try
        {
            if (excelFile == null || excelFile.Length == 0)
                throw new ArgumentException("Excel file is empty or null");

            var dataTable = ReadExcelToDataTable(excelFile);
            var xmlFilePath = Path.Combine(Directory.GetCurrentDirectory(), "output.xml");
            dataTable.WriteXml(xmlFilePath);
            return await File.ReadAllBytesAsync(xmlFilePath);
        }
        catch (Exception ex)
        {
            throw new Exception("Error converting Excel to XML: " + ex.Message);
        }
    }

    private DataTable ReadExcelToDataTable(IFormFile excelFile)
    {
        using (var stream = excelFile.OpenReadStream())
        {
            using (var workbook = new XLWorkbook(stream))
            {
                return workbook.Worksheet(1).RangeUsed().AsTable().AsNativeDataTable();
            }
        }
    }
}