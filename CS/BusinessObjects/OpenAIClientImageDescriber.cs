﻿using Azure;
using Azure.AI.OpenAI;
using DevExpress.Drawing;
using DevExpress.Office.Utils;

namespace RichEditOpenAIWebApi.BusinessObjects {
    class OpenAIClientImageHelper {
        OpenAIClient client;
        internal OpenAIClientImageHelper(string openAIApiKey) {
            client = new OpenAIClient(openAIApiKey, new OpenAIClientOptions());
        }
        string ConvertDXImageToBase64String(DXImage image) {
            using (MemoryStream stream = new MemoryStream()) {
                image.Save(stream, DXImageFormat.Png); 
                byte[] imageBytes = stream.ToArray();
                return Convert.ToBase64String(imageBytes);
            }
        }
        internal async Task<string> DescribeImageAsync(OfficeImage image) {
            string base64Content = ConvertDXImageToBase64String(image.DXImage);
            string imageContentType = OfficeImage.GetContentType(OfficeImageFormat.Png);
            return await GetImageDescription($"data:{imageContentType};base64,{base64Content}");
        }
        internal async Task<string> GetImageDescription(string uriString) {
            ChatCompletionsOptions chatCompletionsOptions = new() {
                DeploymentName = "gpt-4-vision-preview",
                Messages =
                {
                    new ChatRequestSystemMessage("You are a helpful assistant that describes images."),
                    new ChatRequestUserMessage(
                        new ChatMessageTextContentItem("Give a description of this image in no more than 10 words"),
                        new ChatMessageImageContentItem(new Uri(uriString))),
                },
                MaxTokens = 300
            };

            Response<ChatCompletions> chatResponse = await client.GetChatCompletionsAsync(chatCompletionsOptions);
            ChatChoice choice = chatResponse.Value.Choices[0];
            return choice.Message.Content;
        }

    }
}
