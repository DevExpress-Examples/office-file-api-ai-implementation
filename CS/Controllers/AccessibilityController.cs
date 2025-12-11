using DevExpress.AIIntegration;
using DevExpress.Office.Utils;
using DevExpress.Spreadsheet;
using DevExpress.XtraRichEdit;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.AI;
using RichEditOpenAIWebApi.BusinessObjects;
using RichEditOpenAIWebApi.BusinessObjects.RichEditOpenAIWebApi.BusinessObjects;
using Swashbuckle.AspNetCore.Annotations;
using System.Net;

namespace RichEditOpenAIWebApi.Controllers
{
    [ApiController]
    [Route("[controller]/[action]")]
    public class AccessibilityController : ControllerBase
    {
        private readonly IChatClient _chatClient;

        public AccessibilityController(IChatClient chatClient)
        {
            _chatClient = chatClient;
        }
        [HttpPost]
        [SwaggerResponse((int)HttpStatusCode.OK, "Download a file", typeof(FileContentResult))]
        public async Task<IActionResult> GenerateImageAltText(IFormFile documentWithImage, [FromQuery] RichEditFormat outputFormat)
        {
            try
            {
                var imageHelper = new ImageHelper(_chatClient);
                using (var wordProcessor = new RichEditDocumentServer())
                {
                    await RichEditHelper.LoadFile(wordProcessor, documentWithImage);

                    wordProcessor.IterateSubDocuments((document) =>
                    {
                        foreach (var shape in document.Shapes)
                        {
                            if (shape.Type == DevExpress.XtraRichEdit.API.Native.ShapeType.Picture && string.IsNullOrEmpty(shape.AltText))
                            {
                                string description = imageHelper.DescribeImageAsync(shape.PictureFormat.Picture).Result;
                                shape.AltText = description;
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
        public async Task<IActionResult> GenerateChartAltText(IFormFile documentWithImage, [FromQuery] SpreadsheetFormat outputFormat)
        {
            try
            {
                var imageHelper = new ImageHelper(_chatClient);
                using (var workbook = new Workbook())
                {
                    await SpreadsheetHelper.LoadWorkbook(workbook, documentWithImage);

                    foreach (var worksheet in workbook.Worksheets)
                    {
                        foreach (var chart in worksheet.Charts)
                        {
                            OfficeImage image = chart.ExportToImage();
                            string description = await imageHelper.DescribeImageAsync(image);
                            chart.AlternativeText = description;
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
                var hyperlinkHelper = new HyperlinkHelper(_chatClient);
                using (var wordProcessor = new RichEditDocumentServer())
                {
                    await RichEditHelper.LoadFile(wordProcessor, documentWithHyperlinks);

                    wordProcessor.IterateSubDocuments(async (document) =>
                    {
                        foreach (var hyperlink in document.Hyperlinks)
                        {
                            if (string.IsNullOrEmpty(hyperlink.ToolTip) || hyperlink.ToolTip == hyperlink.NavigateUri)
                            {
                                hyperlink.ToolTip = await hyperlinkHelper.DescribeHyperlinkAsync(hyperlink.NavigateUri);
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
                var hyperlinkHelper = new HyperlinkHelper(_chatClient);
                using (var workbook = new Workbook())
                {
                    await SpreadsheetHelper.LoadWorkbook(workbook, documentWithHyperlinks);

                    foreach (var worksheet in workbook.Worksheets)
                    {
                        foreach (var hyperlink in worksheet.Hyperlinks)
                        {
                            if (hyperlink.IsExternal && (string.IsNullOrEmpty(hyperlink.TooltipText) || hyperlink.TooltipText == hyperlink.Uri))
                                hyperlink.TooltipText = await hyperlinkHelper.DescribeHyperlinkAsync(hyperlink.Uri);
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