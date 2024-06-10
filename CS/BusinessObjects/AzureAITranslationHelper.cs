using Azure;
using Azure.AI.Translation.Text;

namespace RichEditOpenAIWebApi.BusinessObjects
{
    public class AzureAITranslationHelper
    {
        TextTranslationClient client;
        internal AzureAITranslationHelper(string key, Uri endpoint, string region = "global")
        {
            AzureKeyCredential azureCredential = new AzureKeyCredential(key);
            client = new TextTranslationClient(azureCredential, endpoint, region);
        }
        internal async Task<string> TranslateText(string text, string sourceLanguage, string targetLanguage)
        {
            Response<IReadOnlyList<TranslatedTextItem>> response = await client.TranslateAsync(targetLanguage, text, sourceLanguage);
            TranslatedTextItem translatedTextItem = response.Value.First();
            return translatedTextItem.Translations[0].Text;
        }
    }
}
