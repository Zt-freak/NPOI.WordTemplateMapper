using NPOI.WordTemplateMapper.XWPF;
using NPOI.XWPF.UserModel;
using NPOI.WordTemplateMapper.Interfaces.XWPF;

namespace NPOI.WordTemplateMapper.Tests.XWPF.DocumentMapperTests
{
    public class MapDocumentTests
    {
        [Fact]
        public void ItShould_MapFullDocument()
        {
            Mock<IXWPFParagraphMapper> paragraphMapperMock = new();
            paragraphMapperMock
                .Setup(p => p.MapParagraph(
                    It.IsAny<XWPFParagraph>(),
                    It.IsAny<IDictionary<string, object>>()
                ))
                .Returns<XWPFParagraph, IDictionary<string, object>>((p, d) => p);

            Mock<IXWPFTableRowMapper> tableRowMapperMock = new();
            tableRowMapperMock
    .Setup(t => t.MapDictionaryToRow(
        It.IsAny<XWPFTableRow>(),
        It.IsAny<IDictionary<string, object>>()
    ))
    .Returns<XWPFTableRow, IDictionary<string, object>>((r, d) => r);
            tableRowMapperMock
                .Setup(t => t.MapEnumerableToRow(
                    It.IsAny<XWPFTableRow>(),
                    It.IsAny<List<Dictionary<string, object>>>()
                ))
                .Returns<XWPFTableRow, List<Dictionary<string, object>>>((r, d) => r);

            Dictionary<string, object> data = new();

            string template = @"TestDocuments/test1.docx";
            using FileStream fileStream = File.OpenRead(template);
            XWPFDocument document = new(fileStream);

            int paragraphCount = document.Paragraphs.Count;
            int footerParagraphCount = document.FooterList.Select(f => f.Paragraphs).Count();
            int headerParagraphCount = document.FooterList.Select(f => f.Paragraphs).Count();
            int rowCount = document.Tables.SelectMany(t => t.Rows).Select(r => r.GetTableCells()).Count();

            XWPFDocumentMapper mapper = new(paragraphMapperMock.Object, tableRowMapperMock.Object);
            mapper.MapDocument(document, data);

            Assert.Equal(paragraphCount, document.Paragraphs.Count);
            Assert.Equal(footerParagraphCount, document.FooterList.Select(f => f.Paragraphs).Count());
            Assert.Equal(headerParagraphCount, document.HeaderList.Select(h => h.Paragraphs).Count());
            paragraphMapperMock.Verify(
                pm => pm.MapParagraph(
                    It.IsAny<XWPFParagraph>(),
                    It.IsAny<IDictionary<string, object>>()),
                    Times.Exactly(
                        document.Paragraphs.Count +
                        document.FooterList.Select(f => f.Paragraphs).Count() +
                        document.HeaderList.Select(h => h.Paragraphs).Count()
                )
            );
            Assert.Equal(rowCount, document.Tables.SelectMany(t => t.Rows).Select(r => r.GetTableCells()).Count());
            tableRowMapperMock.Verify(
                tm => tm.MapDictionaryToRow(It.IsAny<XWPFTableRow>(), It.IsAny<IDictionary<string, object>>()),
                    Times.Exactly(rowCount)
            );
        }
    }
}
