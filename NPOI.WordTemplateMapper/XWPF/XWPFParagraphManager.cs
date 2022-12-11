using NPOI.WordTemplateMapper.Core;
using NPOI.WordTemplateMapper.Core.XWPF;
using NPOI.XWPF.UserModel;
using System.Text.RegularExpressions;

namespace NPOI.WordTemplateMapper.XWPF
{
    public class XWPFParagraphManager : IXWPFParagraphManager
    {
        private static readonly string _alphaNumericSelectorRegex = @"[a-zA-Z0-9.\s\[\]]+";
        private static readonly string _arrayBracketsRegex = @"\[([0-9]+)\]";

        private readonly IKeyValuePairManager _keyValuePairManager;
        private readonly IObjectManager _objectManager;
        public XWPFParagraphManager(IKeyValuePairManager? keyValuePairManager = null, IObjectManager? objectManager = null)
        {
            if (objectManager != null)
                _objectManager = objectManager;
            else
                _objectManager = new ObjectManager();

            if (keyValuePairManager != null)
                _keyValuePairManager = keyValuePairManager;
            else
                _keyValuePairManager = new KeyValuePairManager(_objectManager);
        }

        public IDictionary<string, object> GetContainedMappings(XWPFParagraph paragraph, IDictionary<string, object> mappingDictionary)
        {
            Dictionary<string, object> subsetDictionary = mappingDictionary.Where(mapping => paragraph.Text.Contains(mapping.Key)).ToDictionary(pair => pair.Key, pair => pair.Value);
            return subsetDictionary;
        }

        public XWPFParagraph MapParagraph(XWPFParagraph paragraph, IDictionary<string, object> mappingDictionary)
        {
            foreach (KeyValuePair<string, object> mapping in mappingDictionary)
            {
                if (mapping.Value is IList<object> mappingList)
                {
                    KeyValuePair<string, IList<object>> listMapping = new(mapping.Key, mappingList);
                    Dictionary<string, object> innerMappingDictionary = _keyValuePairManager.ToIndexDictionary(listMapping);
                    MapParagraph(paragraph, innerMappingDictionary);
                }

                bool keepMapping = true;
                do
                {
                    KeyValuePair<string, string> mappedValue = GetMappedValue(paragraph, mapping);
                    string oldText = $"{paragraph.Text}";
                    string newText = paragraph.Text.Replace(mappedValue.Key, mappedValue.Value);

                    // Workaround for malfunctioning ReplaceText from NPOI
                    if (!string.IsNullOrWhiteSpace(newText))
                        paragraph.ReplaceText(paragraph.Text, newText);

                    if (oldText == paragraph.Text)
                        keepMapping = false;
                }
                while (keepMapping);
            }

            return paragraph;
        }

        private KeyValuePair<string, string> GetMappedValue(XWPFParagraph paragraph, KeyValuePair<string, object> mappingToEvaluate)
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

                MatchCollection paragraphTextMatches = Regex.Matches(input: paragraph.Text, pattern: _alphaNumericSelectorRegex);
                if (paragraphTextMatches.Any(m => m.Value.Contains(alphaNumericMappingKey)))
                {
                    KeyValuePair<string, IList<object>> enumerableMappingToEvaluate = new(mappingToEvaluate.Key, mappingList);
                    return MapEnumerableFromPair(paragraph, enumerableMappingToEvaluate, paragraphTextMatches);
                }

                return new(mappingToEvaluate.Key, string.Empty);
            }

            return MapDictionaryFromPair(paragraph, mappingToEvaluate);
        }

        private KeyValuePair<string, string> MapDictionaryFromPair(XWPFParagraph paragraph, KeyValuePair<string, object> mappingToEvaluate)
        {
            Dictionary<string, object> mappingDictionary = _objectManager.ToDictionary(mappingToEvaluate.Value, mappingToEvaluate.Key);
            foreach (KeyValuePair<string, object> mappingPair in mappingDictionary)
            {
                MatchCollection MappingKeyMatches = Regex.Matches(input: mappingPair.Key, pattern: _alphaNumericSelectorRegex);
                string alphaNumericMappingKey = string.Join(string.Empty, from Match match in MappingKeyMatches select match.Value);

                MatchCollection paragraphTextMatches = Regex.Matches(input: paragraph.Text, pattern: _alphaNumericSelectorRegex);
                string alphaNumericParagraphText = string.Join(string.Empty, from Match match in paragraphTextMatches select match.Value);

                if (alphaNumericParagraphText.Contains(alphaNumericMappingKey))
                    return GetMappedValue(paragraph, mappingPair);
            }
            return new(mappingToEvaluate.Key, mappingToEvaluate.Value.ToString()!);
        }

        private KeyValuePair<string, string> MapEnumerableFromPair(XWPFParagraph paragraph, KeyValuePair<string, IList<object>> mappingToEvaluate, MatchCollection paragraphTextMatches)
        {
            Match[] enumerableMatches = paragraphTextMatches.Where(m => Regex.IsMatch(m.Value, _arrayBracketsRegex)).ToArray();

            Dictionary<string, object> mappingDictionary = _keyValuePairManager.ToIndexDictionary(mappingToEvaluate, mappingToEvaluate.Key);
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

                if (paragraph.Text.Contains(mappingPair.Key))
                    return GetMappedValue(paragraph, mappingPair);
            }
            return new(mappingToEvaluate.Key, string.Empty);
        }
    }
}
