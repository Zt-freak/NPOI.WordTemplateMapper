using NPOI.XWPF.UserModel;

namespace NPOI.WordTemplateMapper.Core.XWPF
{
    public interface IXWPFTableRowManager
    {
        public List<XWPFParagraph> GetParagraphsWithMappings(XWPFTableRow row, IDictionary<string, object> mappingDictionary);
        public List<Dictionary<string, object>> GetMappingList(XWPFTableRow row, KeyValuePair<string, IEnumerable<object>>? mappingPair);
        public XWPFTableRow MapDictionaryToRow(XWPFTableRow row, IDictionary<string, object> mappingDictionary);
        public XWPFTableRow MapEnumerableToRow(XWPFTableRow row, List<Dictionary<string, object>> mappingList);
    }
}
