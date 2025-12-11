using DevExpress.AIIntegration;
using DevExpress.AIIntegration.Docs;
using DevExpress.Docs.Presentation;
using DevExpress.Pdf;
using DevExpress.XtraRichEdit;
using Microsoft.AspNetCore.Mvc;
using RichEditOpenAIWebApi.BusinessObjects;
using Swashbuckle.AspNetCore.Annotations;
using System.Net;

namespace RichEditOpenAIWebApi.Controllers
{
    [ApiController]
    [Route("[controller]/[action]")]
    public class SummarizeController : ControllerBase
    {
        private readonly IAIDocProcessingService docProcessingService;

        public SummarizeController(IAIDocProcessingService docService)
        {
            docProcessingService = docService;
        }

        [HttpPost]
        [SwaggerResponse((int)HttpStatusCode.OK, "Download a file", typeof(FileContentResult))]
        public async Task<IActionResult> SummarizeWordDocument(IFormFile document, [FromQuery] SummarizationMode summarizationMode, [FromQuery] RichEditDocumentPart part)
        {
            try
            {
                using (var wordProcessor = new RichEditDocumentServer())
                {
                    await RichEditHelper.LoadFile(wordProcessor, document);
                    string summaryText = string.Empty;

                    switch (part)
                    {
                        case RichEditDocumentPart.FirstPage:
                            var fixedRange = wordProcessor.DocumentLayout.GetPage(0).MainContentRange;
                            var pageRange = wordProcessor.Document.CreateRange(fixedRange.Start, fixedRange.Length);
                            summaryText = await docProcessingService.SummarizeAsync(pageRange, summarizationMode);
                            break;
                        case RichEditDocumentPart.FirstSection:
                            var sectionRange = wordProcessor.Document.Sections[0].Range;
                            summaryText = await docProcessingService.SummarizeAsync(sectionRange, summarizationMode);
                            break;
                        case RichEditDocumentPart.WholeDocument:
                            summaryText = await docProcessingService.SummarizeAsync(wordProcessor, summarizationMode);
                            break;
                    }

                    return Content(summaryText ?? string.Empty, "text/plain");
                }
            }
            catch (Exception e)

            {
                return StatusCode(500, e.Message + Environment.NewLine + e.StackTrace);
            }
        }

        [HttpPost]
        [SwaggerResponse((int)HttpStatusCode.OK, "Download a file", typeof(FileContentResult))]
        public async Task<IActionResult> SummarizePdfDocument(IFormFile file, [FromQuery] SummarizationMode summarizationMode, [FromQuery] PdfPart pdfPart)
        {
            string summaryText = string.Empty;

            try
            {
                using (var pdfDocumentProcessor = new PdfDocumentProcessor())
                {
                    await PdfHelper.LoadPdfDocument(pdfDocumentProcessor, file);
                    switch (pdfPart)
                    {
                        case PdfPart.FirstPage:
                            var firstPageBox = pdfDocumentProcessor.Document.Pages[0].CropBox;
                            PdfDocumentPosition pagePosition1 = new PdfDocumentPosition(1, firstPageBox.TopLeft);
                            PdfDocumentPosition pagePosition2 = new PdfDocumentPosition(1, firstPageBox.BottomRight);
                            var pageArea = PdfDocumentArea.Create(pagePosition1, pagePosition2);
                            summaryText = await docProcessingService.SummarizeAsync(pdfDocumentProcessor, pageArea, summarizationMode);
                            break;
                        case PdfPart.WholeDocument:
                            summaryText = await docProcessingService.SummarizeAsync(pdfDocumentProcessor, summarizationMode);
                            break;
                        default:
                            summaryText = string.Empty; // Ensure summaryText is assigned for all cases
                            break;
                    }

                }
                return Content(summaryText ?? string.Empty, "text/plain");
            }
            catch (Exception e)
            {
                return StatusCode(500, e.Message + Environment.NewLine + e.StackTrace);
            }
        }

        [HttpPost]
        [SwaggerResponse((int)HttpStatusCode.OK, "Download a file", typeof(FileContentResult))]
        public async Task<IActionResult> SummarizePresentation(IFormFile file, [FromQuery] SummarizationMode summarizationMode, [FromQuery] PresentationPart presentationPart)
        {
            string summaryText = string.Empty;
            try
            {
                Presentation presentation = await PresentationHelper.LoadPresentation(file);

                switch (presentationPart)
                {
                    case PresentationPart.FirstSlide:
                        var slide = presentation.Slides[0];
                        summaryText = await docProcessingService.SummarizeAsync(slide, summarizationMode);
                        break;
                    case PresentationPart.WholePresentation:
                        summaryText = await docProcessingService.SummarizeAsync(presentation, summarizationMode);
                        break;
                    default:
                        break;
                }


                return Content(summaryText ?? string.Empty, "text/plain");
            }
            catch (Exception e)
            {
                return StatusCode(500, e.Message + Environment.NewLine + e.StackTrace);
            }
        }


    }
}
