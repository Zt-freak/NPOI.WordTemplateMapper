using System.Text.RegularExpressions;

namespace NPOI.WordTemplateMapper.Extensions
{
    internal static class KeyValuePairExtensions
    {
        private static readonly string alphaNumericSelectorRegex = @"[a-zA-Z0-9.\s\[\]]+";

        internal static List<Dictionary<string, object>> ToList(this KeyValuePair<string, IEnumerable<object>>? @this)
        {
            if (@this == null)
                return new();

            List<Dictionary<string, object>> dictionaryList = new();
            KeyValuePair<string, IEnumerable<object>> mappingPair = (KeyValuePair<string, IEnumerable<object>>)@this;

            foreach (object mappingObject in mappingPair.Value)
            {
                Dictionary<string, object> mappingDictionary = mappingObject.ToDictionary(mappingPair.Key);
                dictionaryList.Add(mappingDictionary);
            }

            return dictionaryList;
        }

        internal static Dictionary<string, object> ToIndexDictionary(this KeyValuePair<string, IList<object>> @this, string? prependKey = null)
        {
            Dictionary<string, object> mappingDictionary = new();

            if (string.IsNullOrEmpty(prependKey))
                prependKey = @this.Key;

            int listItemIndex = 0;
            foreach (object listItem in @this.Value)
            {
                MatchCollection mappingPairKeyMatches = Regex.Matches(input: prependKey, pattern: alphaNumericSelectorRegex);
                string mappingPairKeyWithoutNonAlphanumeric = string.Join(string.Empty, from Match match in mappingPairKeyMatches select match.Value);

                string mappingKey = prependKey.Replace(mappingPairKeyWithoutNonAlphanumeric, $"{mappingPairKeyWithoutNonAlphanumeric}[{listItemIndex}]");

                mappingDictionary.Add(mappingKey, listItem);
                listItemIndex++;
            }
            return mappingDictionary;
        }
    }
}
