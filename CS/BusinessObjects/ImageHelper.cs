using Azure;
using Azure.AI.OpenAI;
using DevExpress.AIIntegration;
using DevExpress.AIIntegration.Docs;
using DevExpress.AIIntegration.Extensions;
using DevExpress.Drawing;
using DevExpress.Office.Utils;
using Microsoft.Extensions.AI;

namespace RichEditOpenAIWebApi.BusinessObjects {
    class ImageHelper {

        private readonly IChatClient _chatClient;
        private readonly AIExtensionsContainerDefault defaultAIExtensionContainer;

        internal ImageHelper(IChatClient chatClient) {
            _chatClient = chatClient;
            defaultAIExtensionContainer = AIExtensionsContainerConsole.CreateDefaultAIExtensionContainer(chatClient);
        }
        string ConvertDXImageToBase64String(DXImage image) {
            using (MemoryStream stream = new MemoryStream()) {
                image.Save(stream, DXImageFormat.Png); 
                byte[] imageBytes = stream.ToArray();
                return Convert.ToBase64String(imageBytes);
            }
        }
        internal async Task<string> DescribeImageAsync(OfficeImage image) {

            var imageByteString = ConvertDXImageToBase64String(image.DXImage);

            var request = new GenerateImageDescriptionRequest(imageByteString);
            var response = await defaultAIExtensionContainer.GenerateImageDescriptionAsync(request);
            return response;
        }
    }
}
