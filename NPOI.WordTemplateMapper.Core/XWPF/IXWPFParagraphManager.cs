using NPOI.XWPF.UserModel;

namespace NPOI.WordTemplateMapper.Core.XWPF
{
    public interface IXWPFParagraphManager
    {
        public IDictionary<string, object> GetContainedMappings(XWPFParagraph paragraph, IDictionary<string, object> mappingDictionary);
        public XWPFParagraph MapParagraph(XWPFParagraph paragraph, IDictionary<string, object> mappingDictionary);
    }
}
