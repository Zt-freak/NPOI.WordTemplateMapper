using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.WordTemplateMapper.Core;
using NPOI.WordTemplateMapper.Core.XWPF;
using NPOI.XWPF.UserModel;

namespace NPOI.WordTemplateMapper.XWPF
{
    public class XWPFTableRowManager : IXWPFTableRowManager
    {
        private readonly IXWPFParagraphManager _paragraphManager;
        private readonly IKeyValuePairManager _keyValuePairManager;
        private readonly IObjectManager _objectManager;

        public XWPFTableRowManager(IXWPFParagraphManager? paragraphManager = null, IKeyValuePairManager? keyValuePairManager = null, IObjectManager? objectManager = null)
        {
            if (objectManager != null)
                _objectManager = objectManager;
            else
                _objectManager = new ObjectManager();

            if (keyValuePairManager != null)
                _keyValuePairManager = keyValuePairManager;
            else
                _keyValuePairManager = new KeyValuePairManager(_objectManager);

            if (paragraphManager != null)
                _paragraphManager = paragraphManager;
            else
                _paragraphManager = new XWPFParagraphManager(_keyValuePairManager, _objectManager);
        }

        public List<XWPFParagraph> GetParagraphsWithMappings(XWPFTableRow row, IDictionary<string, object> mappingDictionary)
        {
            List<XWPFParagraph> rowParagraphs = new();
            row.GetTableCells().ForEach(
                tableCell =>
                {
                    rowParagraphs.AddRange(tableCell.Paragraphs.Where(p => _paragraphManager.GetContainedMappings(p, mappingDictionary).Any()));
                }
            );
            return rowParagraphs;
        }

        public List<Dictionary<string, object>> GetMappingList(XWPFTableRow row, KeyValuePair<string, IEnumerable<object>>? mappingPair)
        {
            if (mappingPair == null)
                return new();

            List<Dictionary<string, object>> mappingList = _keyValuePairManager.ToList(mappingPair);
            if (GetParagraphsWithMappings(row, mappingList.First()).Any())
                return new();

            return mappingList;
        }

        public XWPFTableRow MapDictionaryToRow(XWPFTableRow row, IDictionary<string, object> mappingDictionary)
        {
            List<XWPFParagraph> paragraphs = new();
            row.GetTableCells().ForEach(
                tableCell => paragraphs.AddRange(tableCell.Paragraphs)
            );

            paragraphs.ForEach(
                paragraph => _paragraphManager.MapParagraph(paragraph, mappingDictionary)
            );

            return row;
        }

        public XWPFTableRow MapEnumerableToRow(XWPFTableRow row, List<Dictionary<string, object>> mappingList)
        {
            XWPFTable parentTable = row.GetTable();

            List<XWPFParagraph> paragraphs = new();
            row.GetTableCells().ForEach(
                tableCell => paragraphs.AddRange(tableCell.Paragraphs)
            );

            List<XWPFTableRow> newRows = new();
            for (int i = 0; i < mappingList.Count; i++)
            {
                CT_Row newCtRow = row.GetCTRow().Copy();
                XWPFTableRow copiedRow = new(newCtRow, row.GetTable());
                newRows.Add(copiedRow);
            }

            int newRowsIndex = 0;
            foreach (Dictionary<string, object> mappingDictionary in mappingList)
            {
                int position = parentTable.Rows.IndexOf(row);

                XWPFTableRow copiedRow = newRows[newRowsIndex];
                newRowsIndex++;

                parentTable.AddRow(copiedRow, position);

                MapDictionaryToRow(copiedRow, mappingDictionary);
            }

            parentTable.RemoveRow(parentTable.Rows.IndexOf(row));

            return row;
        }
    }
}
