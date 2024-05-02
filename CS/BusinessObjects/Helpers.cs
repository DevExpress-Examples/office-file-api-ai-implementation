using DevExpress.Spreadsheet;
using DevExpress.XtraRichEdit;
using RichEditDocumentFormat = DevExpress.XtraRichEdit.DocumentFormat;
using RichEditEncryptionSettings = DevExpress.XtraRichEdit.API.Native.EncryptionSettings;
using SpreadsheetDocumentFormat = DevExpress.Spreadsheet.DocumentFormat;
using SpreadsheetEncryptionSettings = DevExpress.Spreadsheet.EncryptionSettings;

namespace RichEditOpenAIWebApi.BusinessObjects
{
    public static class RichEditHelper
    {
        public static async Task LoadFile(RichEditDocumentServer server, IFormFile file)
        {
            using (var stream = new MemoryStream())
            {
                await file.CopyToAsync(stream);
                stream.Seek(0, SeekOrigin.Begin);
                server.LoadDocument(stream);
            }
        }
        public static Stream SaveDocument(RichEditDocumentServer server, RichEditFormat outputFormat, RichEditEncryptionSettings? encryptionSettings = null)
        {
            MemoryStream resultStream = new MemoryStream();
            if (outputFormat == RichEditFormat.Pdf)
                server.ExportToPdf(resultStream);
            else
            {
                RichEditDocumentFormat documentFormat = new RichEditDocumentFormat((int)outputFormat);
                if (documentFormat == RichEditDocumentFormat.Html)
                    server.Options.Export.Html.EmbedImages = true;
                if (encryptionSettings == null)
                    server.SaveDocument(resultStream, documentFormat);
                else
                    server.SaveDocument(resultStream, documentFormat, encryptionSettings);
            }
            resultStream.Seek(0, SeekOrigin.Begin);
            return resultStream;
        }
        public static string GetContentType(RichEditFormat documentFormat)
        {
            switch (documentFormat)
            {
                case RichEditFormat.Doc:
                case RichEditFormat.Dot:
                    return "application/msword";
                case RichEditFormat.Docm:
                case RichEditFormat.Docx:
                case RichEditFormat.Dotm:
                case RichEditFormat.Dotx:
                    return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                case RichEditFormat.ePub:
                    return "application/epub+zip";
                case RichEditFormat.Mht:
                case RichEditFormat.Html:
                    return "text/html";
                case RichEditFormat.Odt:
                    return "application/vnd.oasis.opendocument.text";
                case RichEditFormat.Txt:
                    return "text/plain";
                case RichEditFormat.Rtf:
                    return "application/rtf";
                case RichEditFormat.Xml:
                    return "application/xml";
                case RichEditFormat.Pdf:
                    return "application/pdf";
                default: return string.Empty;
            }
        }
        public static async Task<string> FileToBase64String(IFormFile file)
        {
            using (var stream = new MemoryStream())
            {
                await file.CopyToAsync(stream);
                stream.Seek(0, SeekOrigin.Begin);
                byte[] imageBytes = stream.ToArray();
                return Convert.ToBase64String(imageBytes);
            }
        }
    }
    public static class SpreadsheetHelper
    {

        public static async Task LoadWorkbook(Workbook workbook, IFormFile file)
        {
            using (var stream = new MemoryStream())
            {
                await file.CopyToAsync(stream);
                stream.Seek(0, SeekOrigin.Begin);
                workbook.LoadDocument(stream);
            }
        }

        public static Stream SaveDocument(Workbook workbook, SpreadsheetFormat outputFormat, SpreadsheetEncryptionSettings? encryptionSettings = null)
        {
            MemoryStream resultStream = new MemoryStream();

            SpreadsheetDocumentFormat documentFormat = new SpreadsheetDocumentFormat((int)outputFormat);
            if (outputFormat == SpreadsheetFormat.Pdf)
                workbook.ExportToPdf(resultStream);
            else
            {

                if (encryptionSettings == null)
                    workbook.SaveDocument(resultStream, documentFormat);
                else
                    workbook.SaveDocument(resultStream, documentFormat, encryptionSettings);
            }

            resultStream.Seek(0, SeekOrigin.Begin);
            return resultStream;
        }
        public static string GetContentType(SpreadsheetFormat documentFormat)
        {
            switch (documentFormat)
            {
                case SpreadsheetFormat.Xls:
                case SpreadsheetFormat.Xlt:
                    return "application/msword";
                case SpreadsheetFormat.Xlsx:
                case SpreadsheetFormat.Xlsb:
                case SpreadsheetFormat.Xlsm:
                case SpreadsheetFormat.Xltm:
                    return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                case SpreadsheetFormat.Html:
                    return "text/html";
                case SpreadsheetFormat.XmlSpreadsheet2003:
                    return "application/xml";
                case SpreadsheetFormat.Pdf:
                    return "application/pdf";
                default: return string.Empty;
            }


        }
    }
}
