using NPOI.WordTemplateMapper.XWPF;
using NPOI.XWPF.UserModel;

namespace NPOI.WordTemplateMapper.Tests.XWPF.TableRowMapperTests;

public class MapEnumerableToRowTests
{
    [Fact]
    public void ItShould_MapListToRow()
    {
        XWPFTable parentTable = new(new(), new XWPFDocument());
        XWPFTableRow row = parentTable.GetRow(0);
        row.GetCell(0).Paragraphs[0].CreateRun().SetText("{{Test}}");

        List<Dictionary<string, object>> data = new()
        {
            new()
            {
                { "{{Test}}", "Yeet"}
            },
            new()
            {
                { "{{Test}}", "Dab"}
            },
            new()
            {
                { "{{Test}}", "Yolo"}
            },
            new()
            {
                { "{{Test}}", "Lmao"}
            },
        };

        XWPFTableRowMapper mapper = new();
        mapper.MapEnumerableToRow(row, data);

        Assert.Equal(data.Count, parentTable.Rows.Count);
        Assert.Equal("Yeet", parentTable.Rows[0].GetCell(0).Paragraphs[0].Text);
        Assert.Equal("Dab", parentTable.Rows[1].GetCell(0).Paragraphs[0].Text);
        Assert.Equal("Yolo", parentTable.Rows[2].GetCell(0).Paragraphs[0].Text);
        Assert.Equal("Lmao", parentTable.Rows[3].GetCell(0).Paragraphs[0].Text);
    }
}
