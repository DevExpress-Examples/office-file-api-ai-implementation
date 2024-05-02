using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

namespace RichEditOpenAIWebApi.BusinessObjects {
    static class RichEditExtension {
        internal static void GenerateAltTextForImages(this IRichEditDocumentServer server, Action<SubDocument> action) {
            Document document = server.Document;
            UpdateSubDocument(document, action);
            foreach (var section in document.Sections) {
                UpdateSectionHeaderFooter(section, HeaderFooterType.Primary, action);
                UpdateSectionHeaderFooter(section, HeaderFooterType.First, action);
                UpdateSectionHeaderFooter(section, HeaderFooterType.Odd, action);
                UpdateSectionHeaderFooter(section, HeaderFooterType.Even, action);
            }
        }
        static void UpdateSectionHeaderFooter(Section section, HeaderFooterType type, Action<SubDocument> action) {
            if (section.HasHeader(type)) {
                SubDocument header = section.BeginUpdateHeader(type);
                UpdateSubDocument(header, action);
                section.EndUpdateHeader(header);
            }
            if (section.HasFooter(type)) {
                SubDocument footer = section.BeginUpdateFooter(type);
                UpdateSubDocument(footer, action);
                section.EndUpdateFooter(footer);
            }
        }
        static void UpdateSubDocument(SubDocument document, Action<SubDocument> action) {
            document.BeginUpdate();
            action(document);
            document.EndUpdate();
        }
    }
}
