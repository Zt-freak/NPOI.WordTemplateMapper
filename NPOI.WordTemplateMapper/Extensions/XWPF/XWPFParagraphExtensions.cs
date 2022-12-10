using NPOI.WordTemplateMapper.Extensions;
using NPOI.XWPF.UserModel;
using System.Text.RegularExpressions;

namespace NPOI.WordTemplateMapper.Extensions.XWPF
{
    public static class XWPFParagraphExtensions
    {
        private static readonly string _alphaNumericSelectorRegex = @"[a-zA-Z0-9.\s\[\]]+";
        private static readonly string _arrayBracketsRegex = @"\[([0-9]+)\]";
        public static IDictionary<string, object> GetContainedMappings(this XWPFParagraph @this, IDictionary<string, object> mappingDictionary)
        {
            Dictionary<string, object> subsetDictionary = mappingDictionary.Where(mapping => @this.Text.Contains(mapping.Key)).ToDictionary(pair => pair.Key, pair => pair.Value);
            return subsetDictionary;
        }

        public static XWPFParagraph MapParagraph(this XWPFParagraph @this, IDictionary<string, object> mappingDictionary)
        {
            foreach (KeyValuePair<string, object> mapping in mappingDictionary)
            {
                if (mapping.Value is IList<object> mappingList)
                {
                    KeyValuePair<string, IList<object>> listMapping = new(mapping.Key, mappingList);
                    Dictionary<string, object> innerMappingDictionary = listMapping.ToIndexDictionary();
                    @this.MapParagraph(innerMappingDictionary);
                }

                bool keepMapping = true;
                do
                {
                    KeyValuePair<string, string> mappedValue = @this.GetMappedValue(mapping);
                    string oldText = $"{@this.Text}";
                    string newText = @this.Text.Replace(mappedValue.Key, mappedValue.Value);

                    // Workaround for malfunctioning ReplaceText from NPOI
                    if (!string.IsNullOrWhiteSpace(newText))
                        @this.ReplaceText(@this.Text, newText);

                    if (oldText == @this.Text)
                        keepMapping = false;
                }
                while (keepMapping);
            }

            return @this;
        }

        private static KeyValuePair<string, string> GetMappedValue(this XWPFParagraph @this, KeyValuePair<string, object> mappingToEvaluate)
        {
            if (mappingToEvaluate.Value == null)
                return new(mappingToEvaluate.Key, string.Empty);

            if (mappingToEvaluate.Value.GetType().IsValueType || mappingToEvaluate.Value.GetType() == typeof(string))
                return new(mappingToEvaluate.Key, mappingToEvaluate.Value.ToString()!);

            if (mappingToEvaluate.Value is IEnumerable<object>)
                return new(mappingToEvaluate.Key, string.Empty);

            if (mappingToEvaluate.Value is IList<object> mappingList)
            {
                MatchCollection MappingKeyMatches = Regex.Matches(input: mappingToEvaluate.Key, pattern: _alphaNumericSelectorRegex);
                string alphaNumericMappingKey = string.Join(string.Empty, from Match match in MappingKeyMatches select match.Value);

                MatchCollection paragraphTextMatches = Regex.Matches(input: @this.Text, pattern: _alphaNumericSelectorRegex);
                if (paragraphTextMatches.Any(m => m.Value.Contains(alphaNumericMappingKey)))
                {
                    KeyValuePair<string, IList<object>> enumerableMappingToEvaluate = new(mappingToEvaluate.Key, mappingList);
                    return @this.MapEnumerableFromPair(enumerableMappingToEvaluate, paragraphTextMatches);
                }

                return new(mappingToEvaluate.Key, string.Empty);
            }

            return @this.MapDictionaryFromPair(mappingToEvaluate);
        }

        private static KeyValuePair<string, string> MapDictionaryFromPair(this XWPFParagraph @this, KeyValuePair<string, object> mappingToEvaluate)
        {
            Dictionary<string, object> mappingDictionary = mappingToEvaluate.Value.ToDictionary(mappingToEvaluate.Key);
            foreach (KeyValuePair<string, object> mappingPair in mappingDictionary)
            {
                MatchCollection MappingKeyMatches = Regex.Matches(input: mappingPair.Key, pattern: _alphaNumericSelectorRegex);
                string alphaNumericMappingKey = string.Join(string.Empty, from Match match in MappingKeyMatches select match.Value);

                MatchCollection paragraphTextMatches = Regex.Matches(input: @this.Text, pattern: _alphaNumericSelectorRegex);
                string alphaNumericParagraphText = string.Join(string.Empty, from Match match in paragraphTextMatches select match.Value);

                if (alphaNumericParagraphText.Contains(alphaNumericMappingKey))
                    return @this.GetMappedValue(mappingPair);
            }
            return new(mappingToEvaluate.Key, mappingToEvaluate.Value.ToString()!);
        }

        private static KeyValuePair<string, string> MapEnumerableFromPair(this XWPFParagraph @this, KeyValuePair<string, IList<object>> mappingToEvaluate, MatchCollection paragraphTextMatches)
        {
            Match[] enumerableMatches = paragraphTextMatches.Where(m => Regex.IsMatch(m.Value, _arrayBracketsRegex)).ToArray();

            Dictionary<string, object> mappingDictionary = mappingToEvaluate.ToIndexDictionary(mappingToEvaluate.Key);
            foreach (Match match in enumerableMatches)
            {
                KeyValuePair<string, object> mappingPair = mappingDictionary.FirstOrDefault(m =>
                {
                    MatchCollection MappingKeyMatches = Regex.Matches(input: m.Key, pattern: _alphaNumericSelectorRegex);
                    string alphaNumericMappingKey = string.Join(string.Empty, from Match match in MappingKeyMatches select match.Value);
                    return alphaNumericMappingKey == match.Value;
                });

                if (mappingPair.Key == null)
                    continue;

                if (@this.Text.Contains(mappingPair.Key))
                    return @this.GetMappedValue(mappingPair);
            }
            return new(mappingToEvaluate.Key, string.Empty);
        }
    }
}
