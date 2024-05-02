# Office File API – Integrate AI to Generate Accessible Descriptions

The following project integrates AI capabilities into a DevExpress-powered Office File API Web API application. OpenAI API is used to generate descriptions for images in Microsoft Word files and for Excel charts.

## Implementation Details

The project uses the [Azure.AI.OpenAI](https://www.nuget.org/packages/Azure.AI.OpenAI/) package which adapts OpenAI's REST APIs so it can be used in non-Azure OpenAI development.

The `OpenAIClientImageHelper` class sends a request to describe an image and obtain a string with a response. The `OpenAIClientImageHelper` class methods are executed within corresponding endpoints.

For Excel files, charts are converted to images to obtain a relevant description.

## Files to Review

* [OpenAIController.cs](./CS/Controllers/OpenAIController.cs)
* [OpenAIClientImageHelper.cs](./CS/BusinessObjects/OpenAIClientImageHelper.cs)
* [Helpers.cs](./CS/BusinessObjects/Helpers.cs)

## Documentation

* [Office File API — Enhance Accessibility in Office Documents (Word & Excel) using OpenAI Models](https://community.devexpress.com/blogs/office/archive/2024/04/18/enhance-accessibility-in-office-documents-word-amp-excel-using-artificial-intelligence-system.aspx)