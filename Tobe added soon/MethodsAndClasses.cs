namespace ExportFile.Tobe_added_soon;
public class MethodsAndClasses
{
    #region I will make full Controller for Word
    //There some mistakes!!!

    //public async Task<IActionResult> ConvertToPdf([FromForm] IFormFile file)
    //{
    //    try
    //    {
    //        using (MemoryStream memoryStream = new MemoryStream())
    //        {
    //            await file.CopyToAsync(memoryStream);
    //            using (WordprocessingDocument doc = WordprocessingDocument.Open(memoryStream, false))
    //            {
    //                using (PdfWriter writer = new PdfWriter("output.pdf"))
    //                {
    //                    using (PdfDocument pdf = new PdfDocument(writer))
    //                    {
    //                        DocumentFormat.OpenXml.Wordprocessing.Document document = new DocumentFormat.OpenXml.Wordprocessing.Document(pdf);
    //                        Body body = doc.MainDocumentPart.Document.Body;
    //                        foreach (var element in body.Elements())
    //                        {
    //                            if (element is Paragraph)
    //                            {
    //                                Paragraph paragraph = (Paragraph)element;
    //                                document.Add(new Paragraph(paragraph.InnerText));
    //                            }
    //                        }
    //                        return ("Word document converted to PDF successfully");
    //                    }
    //                }
    //            }
    //        }
    //    }
    //    catch (Exception ex)
    //    {
    //        throw new Exception("Error converting Word to PDF: " + ex.Message);
    //    }
    //}
    #endregion

    #region Make full Controller for Pdf
    //Soon
    #endregion

    #region And some features will be added to develop this project
    //Soon
    #endregion
}