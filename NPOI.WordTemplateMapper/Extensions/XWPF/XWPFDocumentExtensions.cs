using NPOI.XWPF.UserModel;

namespace NPOI.WordTemplateMapper.Extensions.XWPF
{
    public static class XWPFDocumentExtensions
    {
        public static XWPFDocument MapDocument(this XWPFDocument @this, IDictionary<string, object> mappingDictionary)
        {
            @this.MapBody(mappingDictionary);
            @this.MapHeader(mappingDictionary);
            @this.MapFooter(mappingDictionary);
            @this.MapTables(mappingDictionary);

            return @this;
        }

        public static XWPFDocument MapBody(this XWPFDocument @this, IDictionary<string, object> mappingDictionary)
        {
            foreach (XWPFParagraph? paragraph in @this.Paragraphs)
                paragraph.MapParagraph(mappingDictionary);

            return @this;
        }

        public static XWPFDocument MapFooter(this XWPFDocument @this, IDictionary<string, object> mappingDictionary)
        {
            foreach (XWPFFooter footer in @this.FooterList)
                foreach (XWPFParagraph? paragraph in footer.Paragraphs)
                    paragraph.MapParagraph(mappingDictionary);

            return @this;
        }

        public static XWPFDocument MapHeader(this XWPFDocument @this, IDictionary<string, object> mappingDictionary)
        {
            foreach (XWPFHeader header in @this.HeaderList)
                foreach (XWPFParagraph? paragraph in header.Paragraphs)
                    paragraph.MapParagraph(mappingDictionary);

            return @this;
        }

        public static XWPFDocument MapTables(this XWPFDocument @this, IDictionary<string, object> mappingDictionary)
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
            return @this;
        }
    }
}
