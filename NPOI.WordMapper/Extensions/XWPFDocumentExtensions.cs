using NPOI.Util;
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

                for(int i = table.Rows.Count - 1; i >= 0; i--)
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
