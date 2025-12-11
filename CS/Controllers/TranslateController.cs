using DevExpress.AIIntegration.Docs;
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

    public class TranslateController : ControllerBase
    {

        private readonly IAIDocProcessingService docProcessingService;

        public TranslateController(IAIDocProcessingService docService)
        {
            docProcessingService = docService;
        }

        [HttpPost]
        [SwaggerResponse((int)HttpStatusCode.OK, "Download a file", typeof(FileContentResult))]
        public async Task<IActionResult> TranslatePresentation(IFormFile file, [FromQuery] TranslationLang translationLanguage, [FromQuery] PresentationFormat outputFormat, PresentationPart part)
        {
            try
            {
                var presentation = await PresentationHelper.LoadPresentation(file);
                var language = SharedHelper.GetTranslationLanguage(translationLanguage);

                switch (part)
                {
                    case PresentationPart.FirstSlide:
                        var slide = presentation.Slides[0];
                        await docProcessingService.TranslateAsync(slide, language);
                        break;
                    case PresentationPart.WholePresentation:
                        await docProcessingService.TranslateAsync(presentation, language);
                        break;
                    default:
                        break;
                }

                Stream result = PresentationHelper.SaveDocument(presentation, outputFormat);
                string contentType = PresentationHelper.GetContentType(outputFormat);
                string outputStringFormat = outputFormat.ToString().ToLower();

                return File(result, contentType, $"result.{outputStringFormat}");


            }
            catch (Exception e)
            {
                return StatusCode(500, e.Message + Environment.NewLine + e.StackTrace);
            }
        }

        [HttpPost]
        [SwaggerResponse((int)HttpStatusCode.OK, "Download a file", typeof(FileContentResult))]
        public async Task<IActionResult> TranslateWordDocument(IFormFile document, [FromQuery] TranslationLang translationLanguage, [FromQuery] RichEditFormat outputFormat, RichEditDocumentPart part)
        {
            try
            {
                using (var wordProcessor = new RichEditDocumentServer())
                {
                    await RichEditHelper.LoadFile(wordProcessor, document);
                    var language = SharedHelper.GetTranslationLanguage(translationLanguage);

                    switch (part)
                    {
                        case RichEditDocumentPart.FirstPage:
                            var fixedRange = wordProcessor.DocumentLayout.GetPage(0).MainContentRange;
                            var pageRange = wordProcessor.Document.CreateRange(fixedRange.Start, fixedRange.Length); // Translate the first page
                            await docProcessingService.TranslateAsync(pageRange, language);
                            break;
                        case RichEditDocumentPart.FirstSection:
                            var sectionRange = wordProcessor.Document.Sections[0].Range;
                            await docProcessingService.TranslateAsync(sectionRange, language);
                            break;
                        case RichEditDocumentPart.WholeDocument:
                            await docProcessingService.TranslateAsync(wordProcessor, language);
                            break;
                        default:
                            break;
                    }

                    await docProcessingService.TranslateAsync(wordProcessor, language);
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
        public async Task<IActionResult> TranslatePdfDocument(IFormFile document, [FromQuery] TranslationLang translationLanguage, [FromQuery] PdfPart part)
        {
            try
            {
                using (var pdfDocumentProcessor = new PdfDocumentProcessor())
                {
                    await PdfHelper.LoadPdfDocument(pdfDocumentProcessor, document);
                    var language = SharedHelper.GetTranslationLanguage(translationLanguage);
                    string translation = string.Empty; // Initialize to avoid CS0165

                    switch (part)
                    {
                        case PdfPart.FirstPage:
                            var firstPageBox = pdfDocumentProcessor.Document.Pages[0].CropBox;
                            PdfDocumentPosition pagePosition1 = new PdfDocumentPosition(1, firstPageBox.TopLeft);
                            PdfDocumentPosition pagePosition2 = new PdfDocumentPosition(1, firstPageBox.BottomRight);
                            var pageArea = PdfDocumentArea.Create(pagePosition1, pagePosition2);
                            translation = await docProcessingService.TranslateAsync(pdfDocumentProcessor, pageArea, language);
                            break;
                        case PdfPart.WholeDocument:
                            translation = await docProcessingService.TranslateAsync(pdfDocumentProcessor, language);
                            break;
                        default:
                            translation = string.Empty; // Ensure translation is assigned for all cases
                            break;
                    }

                    // Insert a new page and add the translated text
                    PdfPage page = pdfDocumentProcessor.InsertNewPage(1, PdfPaperSize.Letter);
                    PdfRectangle pageSize = page.CropBox;
                    PdfHelper.AddContentToPage(pdfDocumentProcessor, page, pageSize, translation);

                    MemoryStream result = new MemoryStream();
                    pdfDocumentProcessor.SaveDocument(result);
                    result.Seek(0, SeekOrigin.Begin);
                    return File(result, "application/pdf", "result.pdf");
                }

            }
            catch (Exception e)
            {
                return StatusCode(500, e.Message + Environment.NewLine + e.StackTrace);
            }
        }

    }
}
