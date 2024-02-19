using ExportFile.Interfaces;
using Microsoft.AspNetCore.Mvc;

namespace ExportFile.Controllers;

[Route("api/[controller]")]
[ApiController]
public class ExcelController(IExcelService excelService) : ControllerBase
{
    private readonly IExcelService _excelService = excelService;

    [HttpPost("ExcelToJson")]
    public async Task<IActionResult> ExcelToJson(IFormFile excelFile)
    {
        try
        {
            var jsonData = await _excelService.ConvertExcelToJson(excelFile);
            return File(jsonData, "application/json", "output.json");
        }
        catch (Exception ex)
        {
            return StatusCode(500, $"Internal server error: {ex.Message}");
        }
    }

    [HttpPost("ExcelToWord")]
    public async Task<IActionResult> ExcelToWord(IFormFile excelFile)
    {
        try
        {
            var wordData = await _excelService.ConvertExcelToWord(excelFile);
            return File(wordData, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "output.docx");
        }
        catch (Exception ex)
        {
            return StatusCode(500, $"Internal server error: {ex.Message}");
        }
    }

    [HttpPost("ExcelToXML")]
    public async Task<IActionResult> ExcelToXml(IFormFile excelFile)
    {
        try
        {
            var xmlData = await _excelService.ConvertExcelToXml(excelFile);
            return File(xmlData, "application/xml", "output.xml");
        }
        catch (Exception ex)
        {
            return StatusCode(500, $"Internal server error: {ex.Message}");
        }
    }

    [HttpPost("ExcelToPdf")]
    public async Task<IActionResult> ExcelToPdf(IFormFile excelFile)
    {
        try
        {
            var pdfData = await _excelService.ConvertExcelToPdf(excelFile);
            return File(pdfData, "application/pdf", "output.pdf");
        }
        catch (Exception ex)
        {
            return StatusCode(500, $"Internal server error: {ex.Message}");
        }
    }

    [HttpPost("ExcelToCSV")]
    public async Task<IActionResult> ExcelToCSV(IFormFile excelFile)
    {
        try
        {
            var csvData = await _excelService.ConvertExcelToCSV(excelFile);
            return File(csvData, "text/csv", "output.csv");
        }
        catch (Exception ex)
        {
            return StatusCode(500, $"Internal server error: {ex.Message}");
        }
    }
}