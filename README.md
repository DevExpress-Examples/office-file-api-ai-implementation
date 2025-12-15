<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/794940392/25.2.3%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T1231021)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
[![](https://img.shields.io/badge/💬_Leave_Feedback-feecdd?style=flat-square)](#does-this-example-address-your-development-requirementsobjectives)
<!-- default badges end -->
# Office File API – Integrate DevExpress AI-powered Extensions

This project integrates AI capabilities into a Web application that handles user documents. DevExpress Office File API and DevExpress AI-powered Extensions work together to implement the following functionality:

* Generate descriptions for images, charts, and hyperlinks in office files (Word and Excel).
* Summarize, translate, and proofread office files (Word, PDF, and PowerPoint).

> [!note]
> DevExpress does not offer a REST API or ship any built-in LLMs/SLMs. Instead, we follow the BYOL ("bring your own license") principle. To use these features, you need to have an active subscription to AI services (e.g., Azure, Open AI, Google Gemini, Mistral AI, etc.) and obtain the REST API endpoint, key, and model deployment name. These variables must be specified at runtime to enable DevExpress AI-powered Extensions in your application.

## Implementation Details

DevExpress AI-powered extensions run inside an [AIExtensionsContainerDefault](https://docs.devexpress.com/CoreLibraries/DevExpress.AIIntegration.AIExtensionsContainerDefault) container that manages registered AI clients.

The [IChatClient](https://learn.microsoft.com/en-us/dotnet/api/microsoft.extensions.ai.ichatclient) interface serves as the central mechanism for language model interaction.  The `AddDevExpressAIConsole` method registers a chat client in the application. 


The [RegisterAIDocProcessingService(AIExtensionsContainerSettings)](https://docs.devexpress.com/OfficeFileAPI/DevExpress.AIIntegration.Docs.AIDocProcessingExtensions.RegisterAIDocProcessingService(DevExpress.AIIntegration.AIExtensionsContainerSettings)) method registers document-processing AI extensions in a dependency injection container.

The table below lists controllers that use DevExpress AI-powered extensions and corresponding registration methods.

| Controller | Description | API |
|--|----|--|
| `AccessibilityController` | Endpoints that generate image, chart, and hyperlink descriptions.<br/> In Excel files, charts are converted to images to obtain relevant descriptions. | [GenerateImageDescriptionAsync](https://docs.devexpress.com/CoreLibraries/DevExpress.AIIntegration.AIIntegration.GenerateImageDescriptionAsync(IAIExtensionsContainer--GenerateImageDescriptionRequest--CancellationToken))<br/>[CustomPromptAsync](https://docs.devexpress.com/CoreLibraries/DevExpress.AIIntegration.AIIntegration.CustomPromptAsync(IAIExtensionsContainer--CustomPromptRequest--CancellationToken)) |
| `SummarizeController` | Endpoints that produce a concise summary for an entire document/presentation or selected parts (slides, pages, sections). | [SummarizeAsync](https://docs.devexpress.com/OfficeFileAPI/DevExpress.AIIntegration.Docs.IAIDocProcessingService.SummarizeAsync.overloads) |
| `ProofreadController` | Endpoints that review grammar, spelling, and style in an entire document/presentation or selected parts (slides, pages, sections).  | [ProofreadAsync](https://docs.devexpress.com/OfficeFileAPI/DevExpress.AIIntegration.Docs.IAIDocProcessingService.ProofreadAsync.overloads) |
| `TranslateController` | Endpoints that translate a document/presentation or its selected parts (slides, pages, sections). | [TranslateAsync](https://docs.devexpress.com/OfficeFileAPI/DevExpress.AIIntegration.Docs.IAIDocProcessingService.TranslateAsync.overloads) |


## Files to Review

* [Program.cs](./CS/Program.cs)
* [AccessibilityController.cs](./CS/Controllers/AccessibilityController.cs)
* [SummarizeController.cs](./CS/Controllers/SummarizeController.cs)
* [ProofreadController.cs](./CS/Controllers/ProofreadController.cs)
* [TranslateController.cs](./CS/Controllers/TranslateController.cs)
* [Helpers.cs](./CS/BusinessObjects/Helpers.cs)

## Documentation

* [AI-powered Extensions for DevExpress Office File API](https://docs.devexpress.com/OfficeFileAPI/405645/ai-powered-extensions)
* [Office File API — Enhance Accessibility in Office Documents (Word & Excel) using OpenAI Models](https://community.devexpress.com/blogs/office/archive/2024/05/08/enhance-accessibility-in-office-documents-word-and-excel-using-artificial-intelligence-system.aspx)
* [Office File API — Enhance Accessibility in Office Documents using OpenAI Models (Part 2)](https://community.devexpress.com/blogs/office/archive/2024/06/03/office-file-api-enhance-accessibility-in-office-documents-word-amp-excel-using-openai-models-part-2.aspx)
<!-- feedback -->
## Does this example address your development requirements/objectives?

[<img src="https://www.devexpress.com/support/examples/i/yes-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=office-file-api-ai-implementation&~~~was_helpful=yes) [<img src="https://www.devexpress.com/support/examples/i/no-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=office-file-api-ai-implementation&~~~was_helpful=no)

(you will be redirected to DevExpress.com to submit your response)
<!-- feedback end -->

