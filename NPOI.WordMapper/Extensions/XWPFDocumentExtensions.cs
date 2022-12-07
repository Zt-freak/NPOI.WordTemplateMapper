using NPOI.XWPF.UserModel;

namespace NPOI.WordMapper.Extensions
{
    public static class XWPFDocumentExtensions
    {
        public static XWPFDocument MapDocument(this XWPFDocument @this, IDictionary<string, object> mappingDictionary)
        {
            @this.MapTables(mappingDictionary);
            return @this;
        }

        public static XWPFDocument MapTables(this XWPFDocument @this, IDictionary<string, object> mappingDictionary)
        {
            foreach (XWPFTable table in @this.Tables)
            {
                KeyValuePair<string, IEnumerable<object>>? mappingObject = null;
                string tableCaption = table.TableCaption;
                if (mappingDictionary.ContainsKey(tableCaption))
                {
                    IEnumerable<object> mappingEnumerable = (IEnumerable<object>)mappingDictionary[tableCaption];
                    mappingObject = new(tableCaption, mappingEnumerable);
                }

                foreach (XWPFTableRow tableRow in table.Rows)
                {
                    tableRow.MapDictionaryToRow(mappingDictionary);

                    if (mappingObject != null)
                        tableRow.MapEnumerableToRow((KeyValuePair<string, IEnumerable<object>>)mappingObject);
                }
            }

            return @this;
        }
    }
}
