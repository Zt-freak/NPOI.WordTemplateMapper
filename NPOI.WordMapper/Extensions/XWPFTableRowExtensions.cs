using NPOI.OpenXmlFormats.Wordprocessing;
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

        public static List<Dictionary<string, object>> GetMappingList(this XWPFTableRow @this, KeyValuePair<string, IEnumerable<object>>? mappingPair)
        {
            if (mappingPair == null)
                return new();

            List<Dictionary<string, object>> mappingList = mappingPair.ToList();
            if (!@this.GetParagraphsWithMappings(mappingList.First()).Any())
                return new();

            return mappingList;
        }

        public static XWPFTableRow MapDictionaryToRow(this XWPFTableRow @this, IDictionary<string, object> mappingDictionary)
        {
            List<XWPFParagraph> paragraphs = new();
            @this.GetTableCells().ForEach(
                tableCell => paragraphs.AddRange(tableCell.Paragraphs)
            );

            paragraphs.ForEach(
                paragraph => paragraph.MapParagraph(mappingDictionary)
            );

            return @this;
        }

        public static XWPFTableRow MapEnumerableToRow(this XWPFTableRow @this, List<Dictionary<string, object>> mappingList)
        {
            XWPFTable parentTable = @this.GetTable();

            List<XWPFParagraph> paragraphs = new();
            @this.GetTableCells().ForEach(
                tableCell => paragraphs.AddRange(tableCell.Paragraphs)
            );

            List<XWPFTableRow> newRows = new();
            for (int i = 0; i < mappingList.Count(); i++)
            {
                CT_Row newCtRow = @this.GetCTRow().Copy();
                XWPFTableRow copiedRow = new(newCtRow, @this.GetTable());
                newRows.Add(copiedRow);
            }

            int newRowsIndex = 0;
            foreach (Dictionary<string, object> mappingDictionary in mappingList)
            {
                int position = parentTable.Rows.IndexOf(@this);

                XWPFTableRow copiedRow = newRows[newRowsIndex];
                newRowsIndex++;

                parentTable.AddRow(copiedRow, position);

                copiedRow.MapDictionaryToRow(mappingDictionary);
            }

            parentTable.RemoveRow(parentTable.Rows.IndexOf(@this));

            return @this;
        }
    }
}
