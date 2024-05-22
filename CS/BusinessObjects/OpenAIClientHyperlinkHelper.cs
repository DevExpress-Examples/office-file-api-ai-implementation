using Azure;
using Azure.AI.OpenAI;
using DevExpress.Drawing;
using DevExpress.Office.Utils;

namespace RichEditOpenAIWebApi.BusinessObjects
{
    public class OpenAIClientHyperlinkHelper
    {
        OpenAIClient client;
        internal OpenAIClientHyperlinkHelper(string openAIApiKey)
        {
            client = new OpenAIClient(openAIApiKey, new OpenAIClientOptions());
        }
        internal async Task<string> DescribeHyperlinkAsync(string link)
        {
            ChatCompletionsOptions chatCompletionsOptions = new()
            {
                DeploymentName = "gpt-4-vision-preview",
                Messages =
                {
                    new ChatRequestSystemMessage("You are a helpful assistant that describes images."),
                    new ChatRequestUserMessage(
                        new ChatMessageTextContentItem("Give a description of this hyperlink URI in 10-20 words"),
                        new ChatMessageTextContentItem(link))
                },
                MaxTokens = 300
            };

            Response<ChatCompletions> chatResponse = await client.GetChatCompletionsAsync(chatCompletionsOptions);
            ChatChoice choice = chatResponse.Value.Choices[0];
            return choice.Message.Content;
        }
    }
}
