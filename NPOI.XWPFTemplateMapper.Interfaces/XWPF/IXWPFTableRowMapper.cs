using NPOI.XWPF.UserModel;

namespace NPOI.XWPFTemplateMapper.Interfaces.XWPF
{
    public interface IXWPFTableRowMapper
    {
        public List<XWPFParagraph> GetParagraphsWithMappings(XWPFTableRow tableRow, IDictionary<string, object> mappingDictionary);
        public List<Dictionary<string, object>> GetMappingList(XWPFTableRow tableRow, KeyValuePair<string, IEnumerable<object>>? mappingPair);
        public XWPFTableRow MapDictionaryToRow(XWPFTableRow tableRow, IDictionary<string, object> mappingDictionary);
        public XWPFTableRow MapEnumerableToRow(XWPFTableRow tableRow, List<Dictionary<string, object>> mappingList);
    }
}
