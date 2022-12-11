using NPOI.WordTemplateMapper.Core.XWPF;
using NPOI.XWPF.UserModel;

namespace NPOI.WordTemplateMapper.XWPF
{
    public class XWPFTemplateMapper
    {
        private readonly XWPFDocument _document;
        private readonly IDictionary<string, object> _mappingDictionary;
        private readonly IXWPFParagraphManager _paragraphManager;

        public XWPFTemplateMapper(XWPFDocument document, IDictionary<string, object> mappingDictionary, IXWPFParagraphManager paragraphManager)
        {
            _document = document;
            _mappingDictionary = mappingDictionary;
            _paragraphManager = paragraphManager;
        }

        public XWPFDocument MapDocument()
        {
            MapBody();
            MapHeader();
            MapFooter();
            MapTables();

            return _document;
        }

        public XWPFDocument MapBody()
        {
            foreach (XWPFParagraph? paragraph in _document.Paragraphs)
                _paragraphManager.MapParagraph(paragraph, _mappingDictionary);

            return _document;
        }

        public XWPFDocument MapFooter()
        {
            foreach (XWPFFooter footer in _document.FooterList)
                foreach (XWPFParagraph? paragraph in footer.Paragraphs)
                    _paragraphManager.MapParagraph(paragraph, _mappingDictionary);

            return _document;
        }

        public XWPFDocument MapHeader()
        {
            foreach (XWPFHeader header in _document.HeaderList)
                foreach (XWPFParagraph? paragraph in header.Paragraphs)
                    _paragraphManager.MapParagraph(paragraph, _mappingDictionary);

            return _document;
        }

        public XWPFDocument MapTables()
        {
            foreach (XWPFTable table in @this.Tables)
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
                    currentRow.MapDictionaryToRow(mappingDictionary);

                    List<Dictionary<string, object>> mappingList = currentRow.GetMappingList(mappingObject);
                    if (mappingList.Any())
                        currentRow.MapEnumerableToRow(mappingList);
                }
            }
            return _document;
        }
    }
}
