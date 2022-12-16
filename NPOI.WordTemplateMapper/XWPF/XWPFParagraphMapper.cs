using NPOI.WordTemplateMapper.Extensions;
using NPOI.XWPF.UserModel;
using NPOI.WordTemplateMapper.Interfaces.XWPF;
using System.Text.RegularExpressions;
using System.Collections;

namespace NPOI.WordTemplateMapper.XWPF;

public class XWPFParagraphMapper : IXWPFParagraphMapper
{
    private static readonly string _alphaNumericSelectorRegex = @"[a-zA-Z0-9.\s\[\]]+";

    public IDictionary<string, object> GetContainedMappings(XWPFParagraph paragraph, IDictionary<string, object> mappingDictionary)
    {
        Dictionary<string, object> subsetDictionary = mappingDictionary.Where(mapping => paragraph.Text.Contains(mapping.Key)).ToDictionary(pair => pair.Key, pair => pair.Value);
        return subsetDictionary;
    }

    public XWPFParagraph MapParagraph(XWPFParagraph paragraph, IDictionary<string, object> mappingDictionary)
    {
        foreach (KeyValuePair<string, object> mapping in mappingDictionary)
        {
            if (mapping.Value is IList mappingList)
            {
                KeyValuePair<string, IList> listMapping = new(mapping.Key, mappingList);
                Dictionary<string, object> innerMappingDictionary = listMapping.ToIndexDictionary();
                MapParagraph(paragraph, innerMappingDictionary);
            }

            bool keepMapping = true;
            do
            {
                KeyValuePair<string, string> mappedValue = GetMappedValue(paragraph, mapping);
                string oldText = paragraph.Text;

                // Workaround for malfunctioning ReplaceText from NPOI
                string newText = paragraph.Text.Replace(mappedValue.Key, mappedValue.Value);
                if (!string.IsNullOrEmpty(paragraph.Text))
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

        return MapDictionaryFromPair(paragraph, mappingToEvaluate);
    }

    private KeyValuePair<string, string> MapDictionaryFromPair(XWPFParagraph paragraph, KeyValuePair<string, object> mappingToEvaluate)
    {
        Dictionary<string, object> mappingDictionary = mappingToEvaluate.Value.ToDictionary(mappingToEvaluate.Key);
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
}
