using NPOI.WordTemplateMapper.Core;
using NPOI.WordTemplateMapper.Core.XWPF;
using NPOI.XWPF.UserModel;

namespace NPOI.WordTemplateMapper.XWPF
{
    public class XWPFTemplateMapper
    {
        private readonly XWPFDocument _document;
        private readonly IDictionary<string, object> _mappingDictionary;
        private readonly IXWPFParagraphManager _paragraphManager;
        private readonly IXWPFTableRowManager _tableRowManager;
        private readonly IKeyValuePairManager _keyValuePairManager;
        private readonly IObjectManager _objectManager;

        public XWPFTemplateMapper(XWPFDocument document, IDictionary<string, object> mappingDictionary, IXWPFParagraphManager? paragraphManager = null, IXWPFTableRowManager? tableRowManager = null, IKeyValuePairManager? keyValuePairManager = null, IObjectManager? objectManager = null)
        {
            _document = document;
            _mappingDictionary = mappingDictionary;

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

            if (tableRowManager != null)
                _tableRowManager = tableRowManager;
            else
                _tableRowManager = new XWPFTableRowManager(_paragraphManager, _keyValuePairManager, _objectManager);
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
            foreach (XWPFTable table in _document.Tables)
            {
                KeyValuePair<string, IEnumerable<object>>? mappingObject = null;
                string tableCaption = table.TableCaption;

                KeyValuePair<string, object> mappingPair = _mappingDictionary.FirstOrDefault(m => tableCaption.Contains(m.Key));
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

                    List<Dictionary<string, object>> mappingList = _tableRowManager.GetMappingList(currentRow, mappingObject);
                    if (mappingList.Any())
                    {
                        _tableRowManager.MapDictionaryToRow(currentRow, _mappingDictionary);
                        _tableRowManager.MapEnumerableToRow(currentRow, mappingList);
                    }
                }
            }
            return _document;
        }
    }
}
