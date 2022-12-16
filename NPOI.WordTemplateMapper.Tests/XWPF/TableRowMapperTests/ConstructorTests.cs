using NPOI.WordTemplateMapper.XWPF;

namespace NPOI.WordTemplateMapper.Tests.XWPF.TableRowMapperTests;

public class ConstructorTests
{
    [Fact]
    public void ItShould_CreateParagraphMapper_WhenNoneAreInjected()
    {
        XWPFTableRowMapper mapper = new();

        Assert.NotNull(mapper.ParagraphMapper);
        Assert.True(mapper.ParagraphMapper is XWPFParagraphMapper);
    }
}
