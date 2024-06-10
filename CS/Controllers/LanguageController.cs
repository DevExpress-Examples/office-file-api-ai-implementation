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
    public class LanguageController : ControllerBase
    {
        // Insert your Azure key and end point for the Language Service
        string languageAzureKey = "";
        Uri languageEndPoint = new Uri("");
        // Insert your Azure key and end point for the Translation Service
        string translationAzureKey = "";
        Uri translationEndPoint = new Uri("");

        [HttpPost]
        [SwaggerResponse((int)HttpStatusCode.OK, "Download a file", typeof(FileContentResult))]
        public async Task<IActionResult> GenerateLanguageSettingsForParagraphs(IFormFile documentWithHyperlinks, [FromQuery] RichEditFormat outputFormat)
        {
            try
            {
                var languageHelper = new AzureAILanguageHelper(languageAzureKey, languageEndPoint);
                var translationHelper = new AzureAITranslationHelper(translationAzureKey, translationEndPoint);
                using (var server = new RichEditDocumentServer())
                {
                    await RichEditHelper.LoadFile(server, documentWithHyperlinks);

                    server.IterateSubDocuments(async (document) =>
                    {
                        foreach (var paragraph in document.Paragraphs)
                        {
                            CharacterProperties cp = document.BeginUpdateCharacters(paragraph.Range);
                            string paragraphText = document.GetText(paragraph.Range);
                            if (cp.Language.Value.Latin == null && !string.IsNullOrWhiteSpace(paragraphText))
                            {
                                CultureInfo? culture = null;
                                string language = languageHelper.DetectTextLanguage(paragraphText).Result;
                                try
                                {
                                    culture = new CultureInfo((language));
                                }
                                catch { }
                                finally
                                {
                                    if (culture != null)
                                    {
                                        cp.Language = new DevExpress.XtraRichEdit.Model.LangInfo(culture, null, null);
                                        if (language != "en")
                                        {
                                            Comment comment = document.Comments.Create(paragraph.Range, "Translator");
                                            SubDocument commentDoc = comment.BeginUpdate();
                                            string translatedText = translationHelper.TranslateText(paragraphText, language, "en").Result;
                                            commentDoc.InsertText(commentDoc.Range.Start, $"Detected Language: {language}\r\nTranslation (en): {translatedText}");
                                            comment.EndUpdate(commentDoc);
                                        }
                                    }
                                }
                            }
                            document.EndUpdateCharacters(cp);
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
    }
}
