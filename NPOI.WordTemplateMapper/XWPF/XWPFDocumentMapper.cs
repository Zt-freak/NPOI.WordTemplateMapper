using NPOI.XWPF.UserModel;
using NPOI.WordTemplateMapper.Interfaces.XWPF;

namespace NPOI.WordTemplateMapper.XWPF
{
    public class XWPFDocumentMapper : IXWPFDocumentMapper
    {
        private readonly IXWPFParagraphMapper _paragraphMapper;
        private readonly IXWPFTableRowMapper _tableRowMapper;
        public IXWPFParagraphMapper ParagraphMapper { get { return _paragraphMapper; } }
        public IXWPFTableRowMapper TableRowMapper { get { return _tableRowMapper; } }

        public XWPFDocumentMapper(IXWPFParagraphMapper? paragraphMapper = null, IXWPFTableRowMapper? tableRowMapper = null)
        {
            if (paragraphMapper == null)
                _paragraphMapper = new XWPFParagraphMapper();
            else
                _paragraphMapper = paragraphMapper;

            if (tableRowMapper == null)
                _tableRowMapper = new XWPFTableRowMapper();
            else
                _tableRowMapper = tableRowMapper;
        }

        public XWPFDocument MapDocument(XWPFDocument document, IDictionary<string, object> mappingDictionary)
        {
            MapBody(document, mappingDictionary);
            MapHeader(document, mappingDictionary);
            MapFooter(document, mappingDictionary);
            MapTables(document, mappingDictionary);

            return document;
        }

        public XWPFDocument MapBody(XWPFDocument document, IDictionary<string, object> mappingDictionary)
        {
            foreach (XWPFParagraph? paragraph in document.Paragraphs)
                _paragraphMapper.MapParagraph(paragraph, mappingDictionary);

            return document;
        }

        public XWPFDocument MapFooter(XWPFDocument document, IDictionary<string, object> mappingDictionary)
        {
            foreach (XWPFFooter footer in document.FooterList)
                foreach (XWPFParagraph? paragraph in footer.Paragraphs)
                    _paragraphMapper.MapParagraph(paragraph, mappingDictionary);

            return document;
        }

        public XWPFDocument MapHeader(XWPFDocument document, IDictionary<string, object> mappingDictionary)
        {
            foreach (XWPFHeader header in document.HeaderList)
                foreach (XWPFParagraph? paragraph in header.Paragraphs)
                    _paragraphMapper.MapParagraph(paragraph, mappingDictionary);

            return document;
        }

        public XWPFDocument MapTables(XWPFDocument document, IDictionary<string, object> mappingDictionary)
        {
            foreach (XWPFTable table in document.Tables)
            {
                KeyValuePair<string, IEnumerable<object>>? mappingObject = null;
                string tableCaption = table.TableCaption;

                KeyValuePair<string, object> mappingPair = mappingDictionary.FirstOrDefault(m => tableCaption.Contains(m.Key));
                if (mappingPair.Value is IEnumerable<object> mappingEnumerable)
                {
                    mappingObject = new(mappingPair.Key, mappingEnumerable);

                    string newCaption = table.TableCaption.Replace(mappingPair.Key, string.Empty);
                    if (!string.IsNullOrWhiteSpace(newCaption))
                        table.TableCaption = newCaption;
                }

                for (int i = table.Rows.Count - 1; i >= 0; i--)
                {
                    XWPFTableRow currentRow = table.Rows[i];
                    _tableRowMapper.MapDictionaryToRow(currentRow, mappingDictionary);

                    if(mappingObject != null)
                    {
                        List<Dictionary<string, object>> mappingList = _tableRowMapper.GetMappingList(currentRow, mappingObject);
                        if (mappingList != null && mappingList.Any())
                            _tableRowMapper.MapEnumerableToRow(currentRow, mappingList);
                    }
                }
            }
            return document;
        }
    }
}
