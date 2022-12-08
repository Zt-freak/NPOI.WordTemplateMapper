namespace NPOI.WordMapper.Extensions
{
    public static class KeyValuePairExtensions
    {
        public static List<Dictionary<string, object>> ToList(this KeyValuePair<string, IEnumerable<object>>? @this)
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
    }
}
