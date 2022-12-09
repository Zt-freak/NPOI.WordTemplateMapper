using NPOI.XWPF.UserModel;
using System.Text.RegularExpressions;

namespace NPOI.WordMapper.Extensions
{
    public static class XWPFParagraphExtensions
    {
        private static readonly string alphaNumericSelectorRegex = @"[a-zA-Z0-9.\s]+";
        public static IDictionary<string, object> GetContainedMappings(this XWPFParagraph @this, IDictionary<string, object> mappingDictionary)
        {
            Dictionary<string, object> subsetDictionary = mappingDictionary.Where(mapping => @this.Text.Contains(mapping.Key)).ToDictionary(pair => pair.Key, pair => pair.Value);
            return subsetDictionary;
        }

        public static XWPFParagraph MapParagraph(this XWPFParagraph @this, IDictionary<string, object> mappingDictionary)
        {
            foreach (KeyValuePair<string, object> mapping in mappingDictionary)
            {
                KeyValuePair<string, string> mappedValue = @this.GetMappedValue(mapping);
                @this.ReplaceText(mappedValue.Key, mappedValue.Value);
            }

            return @this;
        }

        private static KeyValuePair<string, string> GetMappedValue(this XWPFParagraph @this, KeyValuePair<string, object> mappingToEvaluate)
        {
            if(mappingToEvaluate.Value == null)
                return new(mappingToEvaluate.Key, string.Empty);

            if(mappingToEvaluate.Value.GetType().IsValueType || mappingToEvaluate.Value.GetType() == typeof(string))
                return new(mappingToEvaluate.Key, mappingToEvaluate.Value.ToString()!);

            if(mappingToEvaluate.Value is IEnumerable<object>)
                return new(mappingToEvaluate.Key, string.Empty);

            Dictionary<string, object> mappingDictionary = mappingToEvaluate.Value.ToDictionary(mappingToEvaluate.Key);
            foreach (KeyValuePair<string, object> mappingPair in mappingDictionary)
            {
                MatchCollection paragraphTextMatches = Regex.Matches(input: @this.Text, pattern: alphaNumericSelectorRegex);
                string alphaNumericParagraphText = string.Join(string.Empty, from Match match in paragraphTextMatches select match.Value);

                MatchCollection MappingKeyMathces = Regex.Matches(input: mappingPair.Key, pattern: alphaNumericSelectorRegex);
                string alphaNumericMappingKey = string.Join(string.Empty, from Match match in MappingKeyMathces select match.Value);

                if (alphaNumericParagraphText.Contains(alphaNumericMappingKey))
                    return GetMappedValue(@this, mappingPair);
            }

            return new(mappingToEvaluate.Key, string.Empty);
        }
    }
}
