using NPOI.XWPF.UserModel;

namespace NPOI.WordTemplateMapper.Extensions.XWPF
{
    public class XWPFTableMapper
    {
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
                paragraph.MapParagraph(mappingDictionary);

            return document;
        }

        public XWPFDocument MapFooter(XWPFDocument document, IDictionary<string, object> mappingDictionary)
        {
            foreach (XWPFFooter footer in document.FooterList)
                foreach (XWPFParagraph? paragraph in footer.Paragraphs)
                    paragraph.MapParagraph(mappingDictionary);

            return document;
        }

        public XWPFDocument MapHeader(XWPFDocument document, IDictionary<string, object> mappingDictionary)
        {
            foreach (XWPFHeader header in document.HeaderList)
                foreach (XWPFParagraph? paragraph in header.Paragraphs)
                    paragraph.MapParagraph(mappingDictionary);

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
                    currentRow.MapDictionaryToRow(mappingDictionary);

                    List<Dictionary<string, object>> mappingList = currentRow.GetMappingList(mappingObject);
                    if (mappingList.Any())
                        currentRow.MapEnumerableToRow(mappingList);
                }
            }
            return document;
        }
    }
}
