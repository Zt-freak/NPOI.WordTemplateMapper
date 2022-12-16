using NPOI.WordTemplateMapper.Interfaces.XWPF;
using NPOI.WordTemplateMapper.XWPF;
using NPOI.XWPF.UserModel;

namespace NPOI.WordTemplateMapper.Tests.XWPF.TableRowMapperTests;

public class GetParagraphsWithMappingsTests
{
    [Fact]
    public void ItShould_GetParagraphsWithMappings()
    {
        Mock<IXWPFParagraphMapper> paragraphMapperMock = new();
        paragraphMapperMock
            .Setup(p => p.GetContainedMappings(
                It.IsAny<XWPFParagraph>(),
                It.IsAny<IDictionary<string, object>>()
            ))
            .Returns<XWPFParagraph, IDictionary<string, object>>((p, d) => d);

        XWPFTableRow row = new(new(), new(new(), new XWPFDocument()));
        row.AddNewTableCell().AddParagraph().CreateRun().SetText("{{Test}}");

        Dictionary<string, object> data = new()
        {
            { "Test", "Yeet"}
        };

        XWPFTableRowMapper mapper = new(paragraphMapperMock.Object);
        List<XWPFParagraph> paragraphsList = mapper.GetParagraphsWithMappings(row, data);

        Assert.Equal(row.GetCell(0).Paragraphs.Count, paragraphsList.Count);
    }
}
