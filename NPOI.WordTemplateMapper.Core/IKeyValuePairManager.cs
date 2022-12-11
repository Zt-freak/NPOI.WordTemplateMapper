namespace NPOI.WordTemplateMapper.Core
{
    public interface IKeyValuePairManager
    {
        public List<Dictionary<string, object>> ToList(KeyValuePair<string, IEnumerable<object>>? pair);
        public Dictionary<string, object> ToIndexDictionary(KeyValuePair<string, IList<object>> pair, string? prependKey = null);
    }
}
