using DevExpress.Office.Utils;
using DevExpress.Spreadsheet;
using DevExpress.XtraRichEdit;
using Microsoft.AspNetCore.Mvc;
using RichEditOpenAIWebApi.BusinessObjects;
using Swashbuckle.AspNetCore.Annotations;
using System.Net;

namespace RichEditOpenAIWebApi.Controllers
{
    [ApiController]
    [Route("[controller]/[action]")]
    public class OpenAIController : ControllerBase
    {
        string openAIApiKey = "";
        [HttpPost]
        [SwaggerResponse((int)HttpStatusCode.OK, "Download a file", typeof(FileContentResult))]
        public async Task<IActionResult> GenerateImageAltText(IFormFile documentWithImage, [FromQuery] RichEditFormat outputFormat)
        {
            try
            {
                var imageDescriber = new OpenAIClientImageHelper(openAIApiKey);
                using (var server = new RichEditDocumentServer())
                {
                    await RichEditHelper.LoadFile(server, documentWithImage);

                    server.GenerateAltTextForImages((document) =>
                    {
                        foreach (var shape in document.Shapes)
                        {
                            if (shape.Type == DevExpress.XtraRichEdit.API.Native.ShapeType.Picture && string.IsNullOrEmpty(shape.AltText))
                                shape.AltText = imageDescriber.DescribeImageAsync(shape.PictureFormat.Picture).Result;
                        }
                    });

                    Stream result = RichEditHelper.SaveDocument(server, outputFormat);
                    string contentType = RichEditHelper.GetContentType(outputFormat);
                    string outputStringFormat = outputFormat.ToString().ToLower();
                    return File(result, contentType, $"result.{outputStringFormat}");
                }
            }
            catch (Exception e)
            {
                return StatusCode(500, e.Message + Environment.NewLine + e.StackTrace);
            }
        }

        [HttpPost]
        [SwaggerResponse((int)HttpStatusCode.OK, "Download a file", typeof(FileContentResult))]
        public async Task<IActionResult> GenerateChartAltText(IFormFile documentWithImage, [FromQuery] SpreadsheetFormat outputFormat)
        {
            try
            {
                var imageDescriber = new OpenAIClientImageHelper(openAIApiKey);
                using (var workbook = new Workbook())
                {
                    await SpreadsheetHelper.LoadWorkbook(workbook, documentWithImage);

                    foreach (var worksheet in workbook.Worksheets)
                    {
                        foreach (var chart in worksheet.Charts)
                        {
                            OfficeImage image = chart.ExportToImage();
                            chart.AlternativeText = await imageDescriber.DescribeImageAsync(image);
                        }
                    }

                    Stream result = SpreadsheetHelper.SaveDocument(workbook, outputFormat);
                    string contentType = SpreadsheetHelper.GetContentType(outputFormat);
                    string outputStringFormat = outputFormat.ToString().ToLower();
                    return File(result, contentType, $"result.{outputStringFormat}");
                }
            }
            catch (Exception e)
            {
                return StatusCode(500, e.Message + Environment.NewLine + e.StackTrace);
            }

        }
    }
}