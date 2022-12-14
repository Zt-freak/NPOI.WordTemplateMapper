using NPOI.WordTemplateMapper.XWPF;

namespace NPOI.WordTemplateMapper.Tests.XWPF.DocumentMapperTests
{
    public class ConstructorTests
    {
        [Fact]
        public void ItShould_CreateMappers_WhenNoneAreInjected()
        {
            XWPFDocumentMapper mapper = new();

            Assert.NotNull(mapper.ParagraphMapper);
            Assert.NotNull(mapper.TableRowMapper);
            Assert.True(mapper.ParagraphMapper is XWPFParagraphMapper);
            Assert.True(mapper.TableRowMapper is XWPFTableRowMapper);
        }
    }
}
