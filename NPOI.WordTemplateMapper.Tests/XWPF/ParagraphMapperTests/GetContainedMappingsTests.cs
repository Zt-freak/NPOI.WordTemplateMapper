using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.WordTemplateMapper.XWPF;
using NPOI.XWPF.UserModel;

namespace NPOI.WordTemplateMapper.Tests.XWPF.ParagraphMapperTests
{
    public class GetContainedMappingsTests
    {
        [Theory]
        [InlineData("", 0)]
        [InlineData("lorem ipsum dolor sit amet", 0)]
        [InlineData("the {{A}} fox jumps over the {{B}} dog", 2)]
        [InlineData("Pa's wijze lynx {{F}} het {{M}} aquaduct", 1)]
        [InlineData("{{A}}{{B}}{{C}}{{D}}{{E}}{{F}}{{G}}", 7)]
        [InlineData("{{A}}{{A}}{{B}}{{B}}", 2)]
        public void ItShould_GetMappings_ContainedInParargraph(string paragraphText, int expecctedMappingCount)
        {
            Dictionary<string, object> data = new()
            {
                { "{{A}}", "lorem ipsum" },
                { "{{B}}", "lorem ipsum" },
                { "{{C}}", "lorem ipsum" },
                { "{{D}}", "lorem ipsum" },
                { "{{E}}", "lorem ipsum" },
                { "{{F}}", "lorem ipsum" },
                { "{{G}}", "lorem ipsum" },
            };

            CT_P ctParagraph = new();
            XWPFParagraph paragraph = new(ctParagraph, new XWPFDocument());
            paragraph.CreateRun().SetText(paragraphText);

            XWPFParagraphMapper mapper = new();

            IDictionary<string, object> containedMappings = mapper.GetContainedMappings(paragraph, data);

            Assert.Equal(expecctedMappingCount, containedMappings.Count);
        }
    }
}
