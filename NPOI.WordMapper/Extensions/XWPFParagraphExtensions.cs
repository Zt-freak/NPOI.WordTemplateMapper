using NPOI.XWPF.UserModel;

namespace NPOI.WordMapper.Extensions
{
    public static class XWPFParagraphExtensions
    {
        public static IDictionary<string, object> GetContainedMappings(this XWPFParagraph @this, IDictionary<string, object> mappingDictionary) =>
            (IDictionary<string, object>)mappingDictionary.Where(mapping => @this.Text.Contains(mapping.Key));

        public static IDictionary<string, IEnumerable<object>> GetContainedEnumerableMappings(this XWPFParagraph @this, IDictionary<string, object> mappingDictionary) =>
            (IDictionary<string, IEnumerable<object>>)@this.GetContainedMappings(mappingDictionary).Where(m => typeof(IEnumerable<object>).IsAssignableFrom(m.Value?.GetType()));

        public static IDictionary<string, object> GetContainedNonEnumerableMappings(this XWPFParagraph @this, IDictionary<string, object> mappingDictionary) =>
            (IDictionary<string, object>)@this.GetContainedMappings(mappingDictionary).Where(m => !typeof(IEnumerable<object>).IsAssignableFrom(m.Value?.GetType()));

        public static XWPFParagraph MapParagraph(this XWPFParagraph @this, IDictionary<string, object> mappingDictionary)
        {
            foreach (KeyValuePair<string, object> mapping in mappingDictionary)
                @this.ReplaceText(mapping.Key, mapping.Value.ToString());

            return @this;
        }

        public static XWPFParagraph MapParagraph(this XWPFParagraph @this, object mappingObject)
        {
            throw new NotImplementedException();
        }
    }
}
