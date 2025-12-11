using DevExpress.AIIntegration;
using DevExpress.AIIntegration.Extensions;
using Microsoft.Extensions.AI;
using OpenAI.Chat;
using System.ClientModel;

namespace RichEditOpenAIWebApi.BusinessObjects
{

    namespace RichEditOpenAIWebApi.BusinessObjects
    {
        public class HyperlinkHelper
        {
            private readonly IChatClient _chatClient;
            private readonly AIExtensionsContainerDefault defaultAIExtensionContainer;

            public HyperlinkHelper(IChatClient chatClient)
            {
                _chatClient = chatClient;
                defaultAIExtensionContainer = AIExtensionsContainerConsole.CreateDefaultAIExtensionContainer(chatClient);
            }

            public async Task<string> DescribeHyperlinkAsync(string link, CancellationToken cancellationToken = default)
            {
                if (string.IsNullOrWhiteSpace(link))
                    return string.Empty;
                var request = new CustomPromptRequest("You describe hyperlinks for accessibility. Describe this hyperlink in 10–20 words:", link);
                string response = await defaultAIExtensionContainer.CustomPromptAsync(request, cancellationToken);

                return response.Trim();
            }
        }
    }

}
