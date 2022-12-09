using System.Reflection;
using System.Text.RegularExpressions;

namespace NPOI.WordMapper.Extensions
{
    public static class ObjectExtensions
    {
        private static readonly string alphaNumericSelectorRegex = @"[a-zA-Z0-9.\s\[\]]+";
        public static Dictionary<string, object> ToDictionary(this object @this, string prependKey)
        {
            Dictionary<string, object> mappingDictionary = new();
            PropertyInfo[] propertiesInfo = @this.GetType().GetProperties();

            foreach (PropertyInfo propertyInfo in propertiesInfo)
            {
                MatchCollection mappingPairKeyMatches = Regex.Matches(input: prependKey, pattern: alphaNumericSelectorRegex);
                string mappingPairKeyWithoutNonAlphanumeric = string.Join(string.Empty, from Match match in mappingPairKeyMatches select match.Value);

                string mappingKey = prependKey.Replace(mappingPairKeyWithoutNonAlphanumeric, $"{mappingPairKeyWithoutNonAlphanumeric}.{propertyInfo.Name}");

                mappingDictionary.Add(mappingKey, propertyInfo.GetValue(@this)!);
            }
            return mappingDictionary;
        }

        public static Dictionary<string, IList<object>> ToIndexDictionary(this object @this, string prependKey)
        {
            throw new NotImplementedException();
        }
    }
}
