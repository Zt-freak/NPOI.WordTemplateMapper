using NPOI.XWPF.UserModel;

namespace NPOI.XWPFTemplateMapper.Interfaces.XWPF
{
    public interface IXWPFDocumentMapper
    {
        public XWPFDocument MapDocument(XWPFDocument document, IDictionary<string, object> mappingDictionary);
        public XWPFDocument MapBody(XWPFDocument document, IDictionary<string, object> mappingDictionary);
        public XWPFDocument MapFooter(XWPFDocument document, IDictionary<string, object> mappingDictionary);
        public XWPFDocument MapHeader(XWPFDocument document, IDictionary<string, object> mappingDictionary);
        public XWPFDocument MapTables(XWPFDocument document, IDictionary<string, object> mappingDictionary);
    }
}
