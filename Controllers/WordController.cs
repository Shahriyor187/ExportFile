using ExportFile.Interfaces;
using Microsoft.AspNetCore.Mvc;

namespace ExportFile.Controllers;

[Route("api/[controller]")]
[ApiController]
public class WordController(IWordService wordService) : ControllerBase
{
    private readonly IWordService _wordService = wordService;
    [HttpPost("WordToJson")]
    public async Task<IActionResult> ConvertToJSON(IFormFile wordFile)
    {
        try
        {
            var jsonBytes = await _wordService.ConvertWordToJson(wordFile);
            return File(jsonBytes, "application/json", "output.json");
        }
        catch (Exception ex)
        {
            return StatusCode(500, $"Internal server error: {ex.Message}");
        }
    }
    [HttpPost("WordToExcel")]
    public async Task<IActionResult> ConvertToExcel(IFormFile wordFile)
    {
        try
        {
            var excelBytes = await _wordService.ConvertWordToExcel(wordFile);
            return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "output.xlsx");
        }
        catch (Exception ex)
        {
            return StatusCode(500, $"Internal server error: {ex.Message}");
        }
    }
}