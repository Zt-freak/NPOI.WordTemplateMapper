using NPOI.WordTemplateMapper.Interfaces.XWPF;
using NPOI.WordTemplateMapper.XWPF;
using NPOI.XWPF.UserModel;

namespace NPOI.WordTemplateMapper.Tests.XWPF.TableRowMapperTests;

public class GetMappingListTests
{
    [Fact]
    public void ItShould_GetMappingList()
    {
        Mock<IXWPFParagraphMapper> paragraphMapperMock = new();
        paragraphMapperMock
            .Setup(p => p.GetContainedMappings(
                It.IsAny<XWPFParagraph>(),
                It.IsAny<IDictionary<string, object>>()
            ))
            .Returns<XWPFParagraph, IDictionary<string, object>>((p, d) => d);

        XWPFTableRow row = new(new(), new(new(), new XWPFDocument()));
        row.AddNewTableCell().AddParagraph().CreateRun().SetText("{{Test.A}}");

        List<object> mappingObject = new()
        {
            new{ A = "meow"},
            new{ A = "bark"}
        };

        KeyValuePair<string, IEnumerable<object>>? mappingPair = new("Test", mappingObject);

        XWPFTableRowMapper mapper = new(paragraphMapperMock.Object);
        List<Dictionary<string, object>> mappingList = mapper.GetMappingList(row, mappingPair);

        Assert.Equal(2, mappingList.Count);
        Assert.Equal("Test.A", mappingList[0].First().Key);
        Assert.Equal("meow", mappingList[0].First().Value);
        Assert.Equal("Test.A", mappingList[1].First().Key);
        Assert.Equal("bark", mappingList[1].First().Value);
    }

    [Fact]
    public void ItShould_ReturnEmptyList_IfKeyvaluePairIsNull()
    {
        Mock<IXWPFParagraphMapper> paragraphMapperMock = new();
        paragraphMapperMock
            .Setup(p => p.GetContainedMappings(
                It.IsAny<XWPFParagraph>(),
                It.IsAny<IDictionary<string, object>>()
            ))
            .Returns<XWPFParagraph, IDictionary<string, object>>((p, d) => d);

        XWPFTableRow row = new(new(), new(new(), new XWPFDocument()));

        KeyValuePair<string, IEnumerable<object>>? mappingPair = null;

        XWPFTableRowMapper mapper = new(paragraphMapperMock.Object);
        List<Dictionary<string, object>> mappingList = mapper.GetMappingList(row, mappingPair);

        Assert.Empty(mappingList);
    }

    [Fact]
    public void ItShould_ReturnEmptyList_IfParagraphHasNoMappings()
    {
        Mock<IXWPFParagraphMapper> paragraphMapperMock = new();
        paragraphMapperMock
            .Setup(p => p.GetContainedMappings(
                It.IsAny<XWPFParagraph>(),
                It.IsAny<IDictionary<string, object>>()
            ))
            .Returns<XWPFParagraph, IDictionary<string, object>>((p, d) => new Dictionary<string, object>());

        XWPFTableRow row = new(new(), new(new(), new XWPFDocument()));
        row.AddNewTableCell().AddParagraph().CreateRun().SetText(string.Empty);

        List<object> mappingObject = new()
        {
            new{ A = "meow"},
            new{ A = "bark"}
        };

        KeyValuePair<string, IEnumerable<object>>? mappingPair = new("Test", mappingObject);

        XWPFTableRowMapper mapper = new(paragraphMapperMock.Object);
        List<Dictionary<string, object>> mappingList = mapper.GetMappingList(row, mappingPair);

        Assert.Empty(mappingList);
    }
}
