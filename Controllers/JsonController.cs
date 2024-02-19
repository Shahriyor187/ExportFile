using ExportFile.Interfaces;
using Microsoft.AspNetCore.Mvc;
using System.Net;

namespace ExportFile.Controllers;

[Route("api/[controller]")]
[ApiController]
public class JsonController(IJsonService jsonService) : ControllerBase
{
    private readonly IJsonService _jsonService = jsonService;
    //There are some errors about jsontopdf
    [HttpPost("JsonToPdf")]
    public async Task<IActionResult> JsonToPDF(IFormFile jsonFile)
    {
        try
        {
            var pdfBytes = await _jsonService.ConvertJsonToPdf(jsonFile);
            return File(pdfBytes, "application/pdf", "output.pdf");
        }
        catch (Exception ex)
        {
            return BadRequest("Error converting JSON to PDF: " + ex.Message);
        }
    }

    [HttpPost("JsonToWord")]
    public async Task<IActionResult> JsonToWord(IFormFile jsonFile)
    {
        try
        {
            var wordBytes = await _jsonService.ConvertJsonToWord(jsonFile);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "output.docx");
        }
        catch (Exception ex)
        {
            return BadRequest("Error converting JSON to Word: " + ex.Message);
        }
    }

    [HttpPost("JsonToExcel")]
    public async Task<IActionResult> JsonToExcel(IFormFile jsonFile)
    {
        try
        {
            var excelBytes = await _jsonService.ConvertJsonToExcel(jsonFile);
            return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "output.xlsx");
        }
        catch (Exception ex)
        {
            return BadRequest("Error converting JSON to Excel: " + ex.Message);
        }
    }

    [HttpPost("JsonToCSV")]
    public async Task<IActionResult> JsonToCsv(IFormFile jsonFile)
    {
        try
        {
            var csvBytes = await _jsonService.ConvertJsonToCsv(jsonFile);
            return File(csvBytes, "text/csv", "output.csv");
        }
        catch (Exception ex)
        {
            return StatusCode((int)HttpStatusCode.InternalServerError, ex.Message);
        }
    }
}