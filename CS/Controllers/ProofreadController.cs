using DevExpress.AIIntegration.Docs;
using DevExpress.XtraRichEdit;
using Microsoft.AspNetCore.Mvc;
using RichEditOpenAIWebApi.BusinessObjects;
using Swashbuckle.AspNetCore.Annotations;
using System.Globalization;
using System.Net;

namespace RichEditOpenAIWebApi.Controllers
{
    [ApiController]
    [Route("[controller]/[action]")]
    public class ProofreadController : ControllerBase
    {
        private readonly IAIDocProcessingService docProcessingService;

        public ProofreadController(IAIDocProcessingService docService)
        {
            docProcessingService = docService;
        }

        [HttpPost]
        [SwaggerResponse((int)HttpStatusCode.OK, "Download a file", typeof(FileContentResult))]
        public async Task<IActionResult> ProofreadWordDocument(IFormFile file, [FromQuery] RichEditFormat outputFormat, [FromQuery] RichEditDocumentPart part)
        {
            try
            {
                using (var wordProcessor = new RichEditDocumentServer())
                {
                    await RichEditHelper.LoadFile(wordProcessor, file);
                    CultureInfo cultureInfo = wordProcessor.Document.DefaultCharacterProperties.Language.Value.Latin;

                    switch (part)
                    {
                        case RichEditDocumentPart.FirstPage:
                            var fixedRange = wordProcessor.DocumentLayout.GetPage(0).MainContentRange;
                            var pageRange = wordProcessor.Document.CreateRange(fixedRange.Start, fixedRange.Length); // Translate the first page
                            await docProcessingService.ProofreadAsync(pageRange, cultureInfo);

                            break;
                        case RichEditDocumentPart.FirstSection:
                            var sectionRange = wordProcessor.Document.Sections[0].Range;
                            await docProcessingService.ProofreadAsync(sectionRange, cultureInfo);

                            break;
                        case RichEditDocumentPart.WholeDocument:
                        default:
                            await docProcessingService.ProofreadAsync(wordProcessor, cultureInfo);
                            break;
                    }


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
        public async Task<IActionResult> ProofreadPresentation(IFormFile file, [FromQuery] PresentationPart part, [FromQuery] PresentationFormat outputFormat)
        {
            try
            {
                var presentation = await PresentationHelper.LoadPresentation(file);
                CultureInfo cultureSettings = presentation.DefaultTextStyle.ParagraphProperties.TextProperties.Language;

                switch (part)
                {
                    case PresentationPart.FirstSlide:
                        var slide = presentation.Slides[0];
                        await docProcessingService.ProofreadAsync(slide, cultureSettings);
                        break;
                    case PresentationPart.WholePresentation:
                        await docProcessingService.ProofreadAsync(presentation, cultureSettings);
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
    }
}
