<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/794940392/23.2.5%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T1231021)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
<!-- default badges end -->
# Office File API – Integrate AI to Generate Accessible Descriptions

The following project integrates AI capabilities into a DevExpress-powered Office File API Web API application. OpenAI API is used to generate descriptions for images, charts and hyperlinks in Microsoft Word and Excel files.

> [!note]
> Before you incorporate this solution in your app, please be sure to read and understand OpenAI's license agreement and terms of use.

## Implementation Details

The project uses the [Azure.AI.OpenAI](https://www.nuget.org/packages/Azure.AI.OpenAI/) package which adapts OpenAI's REST APIs so it can be used in non-Azure OpenAI development.

The `OpenAIClientImageHelper` class sends a request to describe an image and obtain a string with a response. The `OpenAIClientHyperlinkHelper` class sends a request to describe an hyperlink and obtain a string with a response. The `OpenAIClientImageHelper.DescribeImageAsync` and `OpenAIClientHyperlinkHelper.DescribeHyperlinkAsync` methods are executed within corresponding endpoints.

For Excel files, charts are converted to images to obtain a relevant description.

## Files to Review

* [OpenAIController.cs](./CS/Controllers/OpenAIController.cs)
* [OpenAIClientImageHelper.cs](./CS/BusinessObjects/OpenAIClientImageHelper.cs)
* [OpenAIClientHyperlinkHelper.cs](./CS/BusinessObjects/OpenAIClientHyperlinkHelper.cs)
* [Helpers.cs](./CS/BusinessObjects/Helpers.cs)

## Documentation

* [Office File API — Enhance Accessibility in Office Documents (Word & Excel) using OpenAI Models](https://community.devexpress.com/blogs/office/archive/2024/05/08/enhance-accessibility-in-office-documents-word-and-excel-using-artificial-intelligence-system.aspx)
