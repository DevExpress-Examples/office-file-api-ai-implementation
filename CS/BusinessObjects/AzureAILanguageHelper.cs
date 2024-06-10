using Azure;
using Azure.AI.TextAnalytics;

namespace RichEditOpenAIWebApi.BusinessObjects
{
    public class AzureAILanguageHelper
    {
        private TextAnalyticsClient client;
        internal AzureAILanguageHelper(string key, Uri endpoint)
        {
            AzureKeyCredential azureCredential = new AzureKeyCredential(key);
            client = new TextAnalyticsClient(endpoint, azureCredential);
        }
        internal async Task<string> DetectTextLanguage(string text)
        {
            DetectedLanguage detectedLanguage = await client.DetectLanguageAsync(text);
            return detectedLanguage.Iso6391Name.Replace('_', '-');
        }
    }
}
