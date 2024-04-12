using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.WordTemplateMapper.XWPF;
using NPOI.XWPF.UserModel;
using System.Globalization;

namespace NPOI.WordTemplateMapper.Tests.XWPF.ParagraphMapperTests;

public class MapParagraphTests
{
    [Theory]
    [InlineData("{{A}} dolor sit amet", "lorem ipsum dolor sit amet")]
    [InlineData("{{B}} jumps over the lazy dog", "the quick brown fox jumps over the lazy dog")]
    [InlineData("{{D}} at all\ntimes.", "The missile\nknows where\nit is at all\ntimes.\n")]
    [InlineData("", "")]
    [InlineData("ABCDEFG", "ABCDEFG")]
    [InlineData("Hey guys\ndid you know\nvaporeon", "Hey guys\ndid you know\nvaporeon")]
    public void ItShould_MapParagraph(string paragraphText, string expecctedText)
    {
        Dictionary<string, object> data = new()
        {
            { "{{A}}", "lorem ipsum" },
            { "{{B}}", "the quick brown fox" },
            { "{{C}}", "het fikse aquaduct" },
            { "{{D}}", "The missile\nknows where\nit is" },
        };

        CT_P ctParagraph = new();
        XWPFParagraph paragraph = new(ctParagraph, new XWPFDocument());
        paragraph.CreateRun().SetText(paragraphText);

        XWPFParagraphMapper mapper = new();
        mapper.MapParagraph(paragraph, data);

        Assert.Equal(expecctedText, paragraph.Text);
    }

    [Theory]
    [InlineData("{{A[0]}} & {{A[1]}}: {{A[4]}}'s Inside Story", "Mario & Luigi: Bowser's Inside Story")]
    [InlineData("{{B[0]}} & {{B[1]}} 3: {{B[4]}} Wars", "Command & Conquer 3: Tiberium Wars")]
    [InlineData("My favourite games are: {{C[0]}}, {{C[1]}} and {{C[2]}}", "My favourite games are: Fortnite, Minecraft and Ace Combat 7")]
    [InlineData("{{D[2].PokeType}} is strong against {{D[1].PokeType}}, which is strong against {{D[0].PokeType}}", "Water is strong against Fire, which is strong against Grass")]
    [InlineData("Did The {{D[3].PokeType}} ever fight {{B[3]}} in WWE?", "Did The Rock ever fight Kane in WWE?")]
    [InlineData("This text has no parameter fields", "This text has no parameter fields")]
    [InlineData("{{E[0][0]}}, {{E[0][1]}}, {{E[0][2]}}, {{E[0][3]}}, {{E[0][4]}}", "Hello, Hallo, Salut, Bonjour, Привет")]
    public void ItShould_MapIListToParagraph(string paragraphText, string expectedText)
    {
        Dictionary<string, object> data = new()
        {
            { "{{A}}", new List<object>{ "Mario", "Luigi", "Donkey Kong", "Kamek", "Bowser" } },
            { "{{B}}", new List<string>{ "Command", "Conquer", "Tiberian", "Kane", "Tiberium" } },
            { "{{C}}", new string[]{ "Fortnite", "Minecraft", "Ace Combat 7" } },
            { "{{D}}", new object[]{ new { PokeType = "Grass" }, new { PokeType = "Fire" }, new { PokeType = "Water" }, new { PokeType = "Rock" }, new { PokeType = "Ice" }, new { PokeType = "Steel" }, new { PokeType = "Ghost" } } },
            { "{{E}}", new object[]{ new string[]{ "Hello", "Hallo", "Salut", "Bonjour", "Привет" } } }
        };

        CT_P ctParagraph = new();
        XWPFParagraph paragraph = new(ctParagraph, new XWPFDocument());
        paragraph.CreateRun().SetText(paragraphText);

        XWPFParagraphMapper mapper = new();
        mapper.MapParagraph(paragraph, data);

        Assert.Equal(expectedText, paragraph.Text);
    }

    [Theory]
    [InlineData("{{A}}", "")]
    [InlineData("{{B}}", "")]
    [InlineData("{{C}}", "   ")]
    [InlineData("test{{A}}test", "testtest")]
    [InlineData("test{{B}}test", "testtest")]
    public void ItShould_MapEmptyToParagraph(string paragraphText, string expectedText)
    {
        Dictionary<string, object> data = new()
        {
            { "{{A}}", null! },
            { "{{B}}", string.Empty },
            { "{{C}}", "   " }
        };

        CT_P ctParagraph = new();
        XWPFParagraph paragraph = new(ctParagraph, new XWPFDocument());
        paragraph.CreateRun().SetText(paragraphText);

        XWPFParagraphMapper mapper = new();
        mapper.MapParagraph(paragraph, data);

        Assert.Equal(expectedText, paragraph.Text);
    }

    [Theory]
    [InlineData("{{A}}", "1")]
    [InlineData("{{B}}", "1.5")]
    [InlineData("{{C}}", "1.5")]
    [InlineData("{{D}}", "d")]
    public void ItShould_MapStructsToParagraph(string paragraphText, string expectedText)
    {
        Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
        Dictionary<string, object> data = new()
        {
            { "{{A}}", 1 },
            { "{{B}}", 1.5 },
            { "{{C}}", new decimal(1.5) },
            { "{{D}}", 'd' },
        };

        CT_P ctParagraph = new();
        XWPFParagraph paragraph = new(ctParagraph, new XWPFDocument());
        paragraph.CreateRun().SetText(paragraphText);

        XWPFParagraphMapper mapper = new();
        mapper.MapParagraph(paragraph, data);

        Assert.Equal(expectedText, paragraph.Text);
    }
}
