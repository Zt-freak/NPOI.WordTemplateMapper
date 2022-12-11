namespace NPOI.WordTemplateMapper.Core.XWPF
{
    public interface IXWPFParagraphManager
    {
        public static IDictionary<string, object> GetContainedMappings(XWPFParagraph @this, IDictionary<string, object> mappingDictionary);

        public static XWPFParagraph MapParagraph(XWPFParagraph paragraph, IDictionary<string, object> mappingDictionary);
    }
}
