namespace NPOI.WordTemplateMapper.Core
{
    public interface IObjectManager
    {
        public Dictionary<string, object> ToDictionary(object mappableObject, string prependKey);
    }
}
