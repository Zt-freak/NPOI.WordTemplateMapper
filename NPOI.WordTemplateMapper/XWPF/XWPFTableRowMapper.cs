using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.WordTemplateMapper.Extensions;
using NPOI.XWPF.UserModel;
using NPOI.WordTemplateMapper.Interfaces.XWPF;

namespace NPOI.WordTemplateMapper.XWPF;

public class XWPFTableRowMapper : IXWPFTableRowMapper
{
    private readonly IXWPFParagraphMapper _paragraphMapper;
    public IXWPFParagraphMapper ParagraphMapper { get { return _paragraphMapper; } }

    public XWPFTableRowMapper(IXWPFParagraphMapper? paragraphMapper = null)
    {
        if (paragraphMapper == null)
            _paragraphMapper = new XWPFParagraphMapper();
        else
            _paragraphMapper = paragraphMapper;
    }

    public List<XWPFParagraph> GetParagraphsWithMappings(XWPFTableRow tableRow, IDictionary<string, object> mappingDictionary)
    {
        List<XWPFParagraph> rowParagraphs = new();
        tableRow.GetTableCells().ForEach(
            tableCell =>
            {
                rowParagraphs.AddRange(tableCell.Paragraphs.Where(p => _paragraphMapper.GetContainedMappings(p, mappingDictionary).Any()));
            }
        );
        return rowParagraphs;
    }

    public List<Dictionary<string, object>> GetMappingList(XWPFTableRow tableRow, KeyValuePair<string, IEnumerable<object>>? mappingPair)
    {
        if (mappingPair == null)
            return new();

        List<Dictionary<string, object>> mappingList = mappingPair.ToList();
        if (!GetParagraphsWithMappings(tableRow, mappingList.First()).Any())
            return new();

        return mappingList;
    }

    public XWPFTableRow MapDictionaryToRow(XWPFTableRow tableRow, IDictionary<string, object> mappingDictionary)
    {
        List<XWPFParagraph> paragraphs = new();
        tableRow.GetTableCells().ForEach(
            tableCell => paragraphs.AddRange(tableCell.Paragraphs)
        );

        paragraphs.ForEach(
            paragraph => _paragraphMapper.MapParagraph(paragraph, mappingDictionary)
        );

        return tableRow;
    }

    public XWPFTableRow MapEnumerableToRow(XWPFTableRow tableRow, List<Dictionary<string, object>> mappingList)
    {
        XWPFTable parentTable = tableRow.GetTable();

        List<XWPFParagraph> paragraphs = new();
        tableRow.GetTableCells().ForEach(
            tableCell => paragraphs.AddRange(tableCell.Paragraphs)
        );

        List<XWPFTableRow> newRows = new();
        for (int i = 0; i < mappingList.Count(); i++)
        {
            CT_Row newCtRow = tableRow.GetCTRow().Copy();
            XWPFTableRow copiedRow = new(newCtRow, tableRow.GetTable());
            newRows.Add(copiedRow);
        }

        int newRowsIndex = 0;
        foreach (Dictionary<string, object> mappingDictionary in mappingList)
        {
            int position = parentTable.Rows.IndexOf(tableRow);

            XWPFTableRow copiedRow = newRows[newRowsIndex];
            newRowsIndex++;

            parentTable.AddRow(copiedRow, position);

            MapDictionaryToRow(copiedRow, mappingDictionary);
        }

        parentTable.RemoveRow(parentTable.Rows.IndexOf(tableRow));

        return tableRow;
    }
}
