using NPOI.WordTemplateMapper.Core;
using System.Text.RegularExpressions;

namespace NPOI.WordTemplateMapper
{
    internal class KeyValuePairManager : IKeyValuePairManager
    {
        private static readonly string alphaNumericSelectorRegex = @"[a-zA-Z0-9.\s\[\]]+";
        private readonly IObjectManager _objectManager;

        public KeyValuePairManager(IObjectManager objectManager)
        {
            if (objectManager != null)
                _objectManager = objectManager;
            else
                _objectManager = new ObjectManager();
        }

        public List<Dictionary<string, object>> ToList(KeyValuePair<string, IEnumerable<object>>? pair)
        {
            if (pair == null)
                return new();

            List<Dictionary<string, object>> dictionaryList = new();
            KeyValuePair<string, IEnumerable<object>> mappingPair = (KeyValuePair<string, IEnumerable<object>>)pair;

            foreach (object mappingObject in mappingPair.Value)
            {
                Dictionary<string, object> mappingDictionary = _objectManager.ToDictionary(mappingObject, mappingPair.Key);
                dictionaryList.Add(mappingDictionary);
            }

            return dictionaryList;
        }

        public Dictionary<string, object> ToIndexDictionary(KeyValuePair<string, IList<object>> pair, string? prependKey = null)
        {
            Dictionary<string, object> mappingDictionary = new();

            if (string.IsNullOrEmpty(prependKey))
                prependKey = pair.Key;

            int listItemIndex = 0;
            foreach (object listItem in pair.Value)
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
