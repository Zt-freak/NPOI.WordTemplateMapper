using NPOI.XWPF.UserModel;

namespace NPOI.WordTemplateMapper.Interfaces.XWPF
{
    public interface IXWPFParagraphMapper
    {
        public IDictionary<string, object> GetContainedMappings(XWPFParagraph paragraph, IDictionary<string, object> mappingDictionary);
        public XWPFParagraph MapParagraph(XWPFParagraph paragraph, IDictionary<string, object> mappingDictionary);
    }
}
