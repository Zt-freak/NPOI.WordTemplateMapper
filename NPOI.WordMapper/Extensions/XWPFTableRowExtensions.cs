using NPOI.XWPF.UserModel;

namespace NPOI.WordMapper.Extensions
{
    public static class XWPFTableRowExtensions
    {
        public static List<XWPFParagraph> GetParagraphsWithMappings(this XWPFTableRow @this, IDictionary<string, object> mappingDictionary)
        {
            List<XWPFParagraph> rowParagraphs = new();
            @this.GetTableCells().ForEach(
                tableCell => {
                    rowParagraphs.AddRange(tableCell.Paragraphs.Where(p => p.GetContainedMappings(mappingDictionary).Any()));
                }
            );
            return rowParagraphs;
        }

        public static List<XWPFParagraph> GetParagraphsWithEnumerableMappings(this XWPFTableRow @this, IDictionary<string, object> mappingDictionary)
        {
            List<XWPFParagraph> rowParagraphs = new();
            @this.GetTableCells().ForEach(
                tableCell => {
                    rowParagraphs.AddRange(tableCell.Paragraphs.Where(p => p.GetContainedEnumerableMappings(mappingDictionary).Any()));
                }
            );
            return rowParagraphs;
        }

        public static XWPFTableRow MapDictionaryToRow(this XWPFTableRow @this, IDictionary<string, object> mappingDictionary)
        {
            throw new NotImplementedException();
        }

        public static XWPFTableRow MapEnumerableToRow(this XWPFTableRow @this, KeyValuePair<string, IEnumerable<object>> mappingObject)
        {
            throw new NotImplementedException();
        }
    }
}
