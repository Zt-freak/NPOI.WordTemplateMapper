using NPOI.WordTemplateMapper.Interfaces.XWPF;
using NPOI.WordTemplateMapper.XWPF;
using NPOI.XWPF.UserModel;

namespace NPOI.WordTemplateMapper.Tests.XWPF.TableRowMapperTests;

public class MapDictionaryToRowTests
{
    [Fact]
    public void ItShould_GetParagraphsWithMappings()
    {
        Mock<IXWPFParagraphMapper> paragraphMapperMock = new();
        paragraphMapperMock
            .Setup(p => p.MapParagraph(
                It.IsAny<XWPFParagraph>(),
                It.IsAny<IDictionary<string, object>>()
            ))
            .Returns<XWPFParagraph, IDictionary<string, object>>((p, d) => p);

        XWPFTableRow row = new(new(), new(new(), new XWPFDocument()));
        row.AddNewTableCell().AddParagraph().CreateRun().SetText("{{Test}}");

        Dictionary<string, object> data = new()
        {
            { "Test", "Yeet"}
        };

        XWPFTableRowMapper mapper = new(paragraphMapperMock.Object);
        mapper.MapDictionaryToRow(row, data);

        Assert.Equal("{{Test}}", row.GetCell(0).Paragraphs[1].Text);
    }
}
