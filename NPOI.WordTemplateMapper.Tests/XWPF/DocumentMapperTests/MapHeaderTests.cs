using NPOI.WordTemplateMapper.XWPF;
using NPOI.XWPF.UserModel;
using NPOI.XWPFTemplateMapper.Interfaces.XWPF;

namespace NPOI.WordTemplateMapper.Tests.XWPF.DocumentMapperTests
{
    public class MapHeaderTests
    {
        [Fact]
        public void ItShould_IterateOverEveryParagraph()
        {
            Mock<IXWPFParagraphMapper> paragraphMapperMock = new();
            paragraphMapperMock
                .Setup(p => p.MapParagraph(
                    It.IsAny<XWPFParagraph>(),
                    It.IsAny<IDictionary<string, object>>()
                ))
                .Returns<XWPFParagraph, IDictionary<string, object>>((p, d) => p);

            Mock<IXWPFTableRowMapper> tableRowMapperMock = new();

            Dictionary<string, object> data = new();

            string template = @"TestDocuments/test1.docx";
            using FileStream fileStream = File.OpenRead(template);
            XWPFDocument document = new(fileStream);

            int paragraphCount = document.HeaderList.Select(h => h.Paragraphs).Count();

            XWPFDocumentMapper mapper = new(paragraphMapperMock.Object, tableRowMapperMock.Object);
            mapper.MapHeader(document, data);

            Assert.Equal(paragraphCount, document.HeaderList.Select(h => h.Paragraphs).Count());
            paragraphMapperMock.Verify(
                pm => pm.MapParagraph(
                    It.IsAny<XWPFParagraph>(),
                    It.IsAny<IDictionary<string, object>>()),
                    Times.Exactly(document.HeaderList.Select(h => h.Paragraphs).Count()
                )
            );
        }
    }
}
