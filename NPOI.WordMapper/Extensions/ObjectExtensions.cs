using System.Reflection;
using System.Text.RegularExpressions;

namespace NPOI.WordMapper.Extensions
{
    public static class ObjectExtensions
    {
        public static Dictionary<string, object> ToDictionary(this object @this, string prependKey)
        {
            Dictionary<string, object> mappingDictionary = new();
            Regex rgx = new("[^a-zA-Z0-9 -]");
            PropertyInfo[] propertiesInfo = @this.GetType().GetProperties();

            foreach (PropertyInfo propertyInfo in propertiesInfo)
            {
                string mappingPairKeyWithoutNonAlphanumeric = rgx.Replace(prependKey, string.Empty);
                string mappingKey = prependKey.Replace(mappingPairKeyWithoutNonAlphanumeric, $"{mappingPairKeyWithoutNonAlphanumeric}.{propertyInfo.Name}");

                mappingDictionary.Add(mappingKey, propertyInfo.GetValue(@this)!);
            }
            return mappingDictionary;
        }
    }
}
