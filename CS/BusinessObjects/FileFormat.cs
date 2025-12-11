namespace RichEditOpenAIWebApi.BusinessObjects
{
    public enum RichEditFormat
    {
        Txt = 1,
        Rtf = 2,
        Html = 3,
        Docx = 4,
        Mht = 5,
        Xml = 6, // WordML
        Odt = 7,
        ePub = 9,
        Doc = 10,
        Pdf = 11,
        Docm = 12,
        Dotx = 13,
        Dotm = 14,
        Dot = 15,
    }

    public enum RichEditDocumentPart
    {
        WholeDocument = 1,
        FirstPage = 2,
        FirstSection = 3,
    }

    public enum SpreadsheetFormat
    {
        Xls = 1,
        Xlsx = 2,
        Html = 6,
        Xlsm = 7,
        Xlt = 8,
        Xltx = 9,
        Xltm = 10,
        Xlsb = 11,
        XmlSpreadsheet2003 = 12,
        Pdf = 13,
    }

    public enum PresentationFormat
    {
        Pptx = 1,
        Ppt = 2,
        Pdf = 3,
    }

    public enum PresentationPart
    {
        WholePresentation = 1,
        FirstSlide = 2,
    }

    public enum PdfPart
    {
        WholeDocument = 1,
        FirstPage = 2,
    }
    public enum TranslationLang
    {
        English = 1,
        Spanish = 2,
        French = 3,
        German = 4,

    }

}
