using DevExpress.Office.Utils;
using DevExpress.Spreadsheet;
using DevExpress.XtraPrinting.Native;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using Microsoft.AspNetCore.Mvc;
using RichEditOpenAIWebApi.BusinessObjects;
using Swashbuckle.AspNetCore.Annotations;
using System.Globalization;
using System.Net;

namespace RichEditOpenAIWebApi.Controllers
{
    [ApiController]
    [Route("[controller]/[action]")]
    public class OpenAIController : ControllerBase
    {
        // Insert your OpenAI key
        string openAIApiKey = "";
        [HttpPost]
        [SwaggerResponse((int)HttpStatusCode.OK, "Download a file", typeof(FileContentResult))]
        public async Task<IActionResult> GenerateImageAltText(IFormFile documentWithImage, [FromQuery] RichEditFormat outputFormat)
        {
            try
            {
                var imageHelper = new OpenAIClientImageHelper(openAIApiKey);
                using (var wordProcessor = new RichEditDocumentServer())
                {
                    await RichEditHelper.LoadFile(wordProcessor, documentWithImage);

                    wordProcessor.IterateSubDocuments((document) =>
                    {
                        foreach (var shape in document.Shapes)
                        {
                            if (shape.Type == DevExpress.XtraRichEdit.API.Native.ShapeType.Picture && string.IsNullOrEmpty(shape.AltText))
                                shape.AltText = imageHelper.DescribeImageAsync(shape.PictureFormat.Picture).Result;
                        }
                    });

                    Stream result = RichEditHelper.SaveDocument(wordProcessor, outputFormat);
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
                var imageHelper = new OpenAIClientImageHelper(openAIApiKey);
                using (var workbook = new Workbook())
                {
                    await SpreadsheetHelper.LoadWorkbook(workbook, documentWithImage);

                    foreach (var worksheet in workbook.Worksheets)
                    {
                        foreach (var chart in worksheet.Charts)
                        {
                            OfficeImage image = chart.ExportToImage();
                            chart.AlternativeText = imageHelper.DescribeImageAsync(image).Result;
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
        [HttpPost]
        [SwaggerResponse((int)HttpStatusCode.OK, "Download a file", typeof(FileContentResult))]
        public async Task<IActionResult> GenerateHyperlinkDescriptionForWord(IFormFile documentWithHyperlinks, [FromQuery] RichEditFormat outputFormat)
        {
            try
            {
                var hyperlinkHelper = new OpenAIClientHyperlinkHelper(openAIApiKey);
                using (var wordProcessor = new RichEditDocumentServer())
                {
                    await RichEditHelper.LoadFile(wordProcessor, documentWithHyperlinks);

                    wordProcessor.IterateSubDocuments(async (document) =>
                    {
                        foreach (var hyperlink in document.Hyperlinks)
                        {
                            if (string.IsNullOrEmpty(hyperlink.ToolTip) || hyperlink.ToolTip == hyperlink.NavigateUri)
                            {
                                hyperlink.ToolTip = hyperlinkHelper.DescribeHyperlinkAsync(hyperlink.NavigateUri).Result;
                            }
                        }
                    });
                    Stream result = RichEditHelper.SaveDocument(wordProcessor, outputFormat);
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
        public async Task<IActionResult> GenerateHyperlinkDescriptionForSpreadsheet(IFormFile documentWithHyperlinks, [FromQuery] SpreadsheetFormat outputFormat)
        {
            try
            {
                var hyperlinkHelper = new OpenAIClientHyperlinkHelper(openAIApiKey);
                using (var workbook = new Workbook())
                {
                    await SpreadsheetHelper.LoadWorkbook(workbook, documentWithHyperlinks);

                    foreach (var worksheet in workbook.Worksheets)
                    {
                        foreach (var hyperlink in worksheet.Hyperlinks)
                        {
                            if(hyperlink.IsExternal && (string.IsNullOrEmpty(hyperlink.TooltipText) || hyperlink.TooltipText == hyperlink.Uri))
                            hyperlink.TooltipText = hyperlinkHelper.DescribeHyperlinkAsync(hyperlink.Uri).Result;
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