using NPOI.WordTemplateMapper.Core;
using System.Reflection;
using System.Text.RegularExpressions;

namespace NPOI.WordTemplateMapper
{
    internal class ObjectManager : IObjectManager
    {
        private static readonly string alphaNumericSelectorRegex = @"[a-zA-Z0-9.\s\[\]]+";

        public Dictionary<string, object> ToDictionary(object mappableObject, string prependKey)
        {
            Dictionary<string, object> mappingDictionary = new();
            PropertyInfo[] propertiesInfo = mappableObject.GetType().GetProperties();

            foreach (PropertyInfo propertyInfo in propertiesInfo)
            {
                MatchCollection mappingPairKeyMatches = Regex.Matches(input: prependKey, pattern: alphaNumericSelectorRegex);
                string mappingPairKeyWithoutNonAlphanumeric = string.Join(string.Empty, from Match match in mappingPairKeyMatches select match.Value);

                string mappingKey = prependKey.Replace(mappingPairKeyWithoutNonAlphanumeric, $"{mappingPairKeyWithoutNonAlphanumeric}.{propertyInfo.Name}");

                mappingDictionary.Add(mappingKey, propertyInfo.GetValue(mappableObject)!);
            }
            return mappingDictionary;
        }
    }
}
