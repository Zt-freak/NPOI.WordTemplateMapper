using NPOI.XWPF.UserModel;

namespace NPOI.WordMapper.Extensions
{
    public static class XWPFParagraphExtensions
    {
        public static IDictionary<string, object> GetContainedMappings(this XWPFParagraph @this, IDictionary<string, object> mappingDictionary)
        {
            Dictionary<string, object> subsetDictionary = mappingDictionary.Where(mapping => @this.Text.Contains(mapping.Key)).ToDictionary(pair => pair.Key, pair => pair.Value);
            return subsetDictionary;
        }

        public static IDictionary<string, IEnumerable<object>> GetContainedEnumerableMappings(this XWPFParagraph @this, IDictionary<string, object> mappingDictionary) =>
            (IDictionary<string, IEnumerable<object>>)@this.GetContainedMappings(mappingDictionary).Where(m => typeof(IEnumerable<object>).IsAssignableFrom(m.Value?.GetType()));

        public static bool UsesMappingDictionary(this XWPFParagraph @this, IDictionary<string, object> mappingDictionary) =>
            @this.GetContainedMappings(mappingDictionary).Any();

        public static XWPFParagraph MapParagraph(this XWPFParagraph @this, IDictionary<string, object> mappingDictionary)
        {
            foreach (KeyValuePair<string, object> mapping in mappingDictionary)
                @this.ReplaceText(mapping.Key, TryGetValue(mapping.Value));

            return @this;
        }

        private static string TryGetValue(object objectToEvaluate)
        {
            if (objectToEvaluate == null)
                return string.Empty;

            if (objectToEvaluate.GetType().IsValueType || objectToEvaluate.GetType() == typeof(string))
                return objectToEvaluate.ToString()!;

            return string.Empty;
        }
    }
}
