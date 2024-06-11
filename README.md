<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/794940392/23.2.5%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T1231021)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
<!-- default badges end -->
# Office File API – Integrate AI

The following project integrates AI capabilities into a DevExpress-powered Office File API Web API application. This project uses the following AI services:
* OpenAI API generates descriptions for images, charts and hyperlinks in Microsoft Word and Excel files.
* Azure AI Language API detects the language for text paragraphs in Word files.
* Azure AI Translator API translates paragraph text to the chosen language in Word files.

> [!note]
> Before you incorporate this solution in your app, please be sure to read and understand terms and conditions of using AI services.

## Implementation Details

The project uses the [Azure.AI.OpenAI](https://www.nuget.org/packages/Azure.AI.OpenAI/), [Azure.AI.TextAnalytics](https://www.nuget.org/packages/Azure.AI.TextAnalytics) and [Azure.AI.Translation.Text](https://www.nuget.org/packages/Azure.AI.Translation.Text) NuGet packages. Azure.AI.OpenAI adapts OpenAI's REST APIs for use in non-Azure OpenAI development. Azure.AI.TextAnalytics and Azure.AI.Translation.Text require an Azure subscription. Once you obtain it, create a [Language resource](https://portal.azure.com/#create/Microsoft.CognitiveServicesTextAnalytics) and a [Translator resource](https://portal.azure.com/#create/Microsoft.CognitiveServicesTextTranslation) (or a single [multi-service resource](https://learn.microsoft.com/en-us/azure/ai-services/multi-service-resource?tabs=windows&pivots=azportal)) in the Azure portal to get your keys and endpoints for client authentification.

`OpenAIController` includes endpoints to generate image, chart and hyperlink descriptions. The `OpenAIClientImageHelper` class sends a request to describe an image and obtains a string with a response. The `OpenAIClientHyperlinkHelper` class sends a request to describe an hyperlink and obtains a string with a response. The `OpenAIClientImageHelper.DescribeImageAsync` and `OpenAIClientHyperlinkHelper.DescribeHyperlinkAsync` methods are executed within the corresponding endpoints.
For Excel files, charts are converted to images to obtain relevant descriptions.

`LanguageController` includes the endpoint to detect the language for text paragraphs and generate paragraph transaltions. The `AzureAILanguageHelper` class sends a request to detect the language of the specified text and returns the language name in the "ISO 693-1" format. The `AzureAITranslationHelper` class sends a request to translate the given text to the specified language and returns the transaled text string. The `AzureAILanguageHelper.DetectTextLanguage` and `AzureAITranslationHelper.TranslateText` methods are executed in the  `GenerateLanguageSettingsForParagraphs` endpoint.

## Files to Review

* [OpenAIController.cs](./CS/Controllers/OpenAIController.cs)
* [LanguageController.cs](./CS/Controllers/LanguageController.cs)
* [OpenAIClientImageHelper.cs](./CS/BusinessObjects/OpenAIClientImageHelper.cs)
* [OpenAIClientHyperlinkHelper.cs](./CS/BusinessObjects/OpenAIClientHyperlinkHelper.cs)
* [AzureAILanguageHelper.cs](./CS/BusinessObjects/AzureAILanguageHelper.cs)
* [AzureAITranslationHelper.cs](./CS/BusinessObjects/AzureAITranslationHelper.cs)
* [Helpers.cs](./CS/BusinessObjects/Helpers.cs)

## Documentation

* [Office File API — Enhance Accessibility in Office Documents (Word & Excel) using OpenAI Models](https://community.devexpress.com/blogs/office/archive/2024/05/08/enhance-accessibility-in-office-documents-word-and-excel-using-artificial-intelligence-system.aspx)
* [Office File API — Enhance Accessibility in Office Documents using OpenAI Models (Part 2)](https://community.devexpress.com/blogs/office/archive/2024/06/03/office-file-api-enhance-accessibility-in-office-documents-word-amp-excel-using-openai-models-part-2.aspx)
