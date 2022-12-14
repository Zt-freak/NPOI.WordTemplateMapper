using NPOI.WordTemplateMapper.XWPF;
using NPOI.XWPF.UserModel;
using NPOI.WordTemplateMapper.Interfaces.XWPF;

namespace NPOI.WordTemplateMapper.Tests.XWPF.DocumentMapperTests
{
    public class MapTablesTests
    {
        [Fact]
        public void ItShould_IterateOverEveryTableRow()
        {
            Mock<IXWPFParagraphMapper> paragraphMapperMock = new();

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

            Dictionary<string, object> data = new()
            {
                { "{{CoolSinger}}", "Rick Astley" }
            };

            string template = @"TestDocuments/test2.docx";
            using FileStream fileStream = File.OpenRead(template);
            XWPFDocument document = new(fileStream);

            int rowCount = document.Tables.SelectMany(t => t.Rows).Select(r => r.GetTableCells()).Count();

            XWPFDocumentMapper mapper = new(paragraphMapperMock.Object, tableRowMapperMock.Object);
            mapper.MapTables(document, data);

            Assert.Equal(rowCount, document.Tables.SelectMany(t => t.Rows).Select(r => r.GetTableCells()).Count());
            tableRowMapperMock.Verify(
                tm => tm.MapDictionaryToRow(It.IsAny<XWPFTableRow>(), It.IsAny<IDictionary<string, object>>()),
                    Times.Exactly(rowCount)
            );
        }

        [Fact]
        public void ItShould_MapEnumerablesToTableRow()
        {
            Mock<IXWPFParagraphMapper> paragraphMapperMock = new();

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
                .Returns<XWPFTableRow, List<Dictionary<string, object>>>((r, d) => r); //GetMappingList
            tableRowMapperMock
                .Setup(t => t.GetMappingList(
                    It.IsAny<XWPFTableRow>(),
                    It.IsAny<KeyValuePair<string, IEnumerable<object>>>()
                ))
                .Returns<XWPFTableRow, KeyValuePair<string, IEnumerable<object>>>((r, d) => new() { new() });

            List<object> theList = new() {
                new { A = "Chirp", B = "Meow", C = "Woof", D = "Moo", E = "Bonjour", F = "Quack"},
                new { A = "Monday", B = "Tuesday", C = "Wednesday", D = "Thursday", E = "Friday", F = "Saturday"},
                new { A = "Red", B = "Green", C = "Blue", D = "Yellow", E = "Magenta", F = "Cyan"}
            };
            Dictionary<string, object> data = new()
            {
                { "{{List}}", theList }
            };

            string template = @"TestDocuments/test3.docx";
            using FileStream fileStream = File.OpenRead(template);
            XWPFDocument document = new(fileStream);

            int rowCount = document.Tables.SelectMany(t => t.Rows).Select(r => r.GetTableCells()).Count();

            XWPFDocumentMapper mapper = new(paragraphMapperMock.Object, tableRowMapperMock.Object);
            mapper.MapTables(document, data);

            tableRowMapperMock.Verify(
                tm => tm.MapEnumerableToRow(It.IsAny<XWPFTableRow>(), It.IsAny<List<Dictionary<string, object>>>()),
                    Times.Exactly(rowCount)
            );
        }

        [Fact]
        public void ItShould_CleanUpTableCaption()
        {
            Mock<IXWPFParagraphMapper> paragraphMapperMock = new();

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
                .Returns<XWPFTableRow, List<Dictionary<string, object>>>((r, d) => r); //GetMappingList
            tableRowMapperMock
                .Setup(t => t.GetMappingList(
                    It.IsAny<XWPFTableRow>(),
                    It.IsAny<KeyValuePair<string, IEnumerable<object>>>()
                ))
                .Returns<XWPFTableRow, KeyValuePair<string, IEnumerable<object>>>((r, d) => new() { new() });

            List<object> theList = new() {
                new { A = "Chirp", B = "Meow", C = "Woof", D = "Moo", E = "Bonjour", F = "Quack"},
                new { A = "Monday", B = "Tuesday", C = "Wednesday", D = "Thursday", E = "Friday", F = "Saturday"},
                new { A = "Red", B = "Green", C = "Blue", D = "Yellow", E = "Magenta", F = "Cyan"}
            };
            Dictionary<string, object> data = new()
            {
                { "{{List}}", theList }
            };

            string template = @"TestDocuments/test3.docx";
            using FileStream fileStream = File.OpenRead(template);
            XWPFDocument document = new(fileStream);

            string oldCaption = document.Tables[0].TableCaption;

            XWPFDocumentMapper mapper = new(paragraphMapperMock.Object, tableRowMapperMock.Object);
            mapper.MapTables(document, data);

            Assert.False(oldCaption == document.Tables[0].TableCaption);
            Assert.DoesNotContain("{{List}}", document.Tables[0].TableCaption);
        }

        [Fact]
        public void ItShould_MapNoEnumerables_IfMappingListIsEmpty()
        {
            Mock<IXWPFParagraphMapper> paragraphMapperMock = new();

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
                .Returns<XWPFTableRow, List<Dictionary<string, object>>>((r, d) => r); //GetMappingList
            tableRowMapperMock
                .Setup(t => t.GetMappingList(
                    It.IsAny<XWPFTableRow>(),
                    It.IsAny<KeyValuePair<string, IEnumerable<object>>>()
                ))
                .Returns<XWPFTableRow, KeyValuePair<string, IEnumerable<object>>>((r, d) => new() { });

            List<object> theList = new() {
                new { A = "Chirp", B = "Meow", C = "Woof", D = "Moo", E = "Bonjour", F = "Quack"},
                new { A = "Monday", B = "Tuesday", C = "Wednesday", D = "Thursday", E = "Friday", F = "Saturday"},
                new { A = "Red", B = "Green", C = "Blue", D = "Yellow", E = "Magenta", F = "Cyan"}
            };
            Dictionary<string, object> data = new()
            {
                { "{{List}}", theList }
            };

            string template = @"TestDocuments/test3.docx";
            using FileStream fileStream = File.OpenRead(template);
            XWPFDocument document = new(fileStream);

            int rowCount = document.Tables.SelectMany(t => t.Rows).Select(r => r.GetTableCells()).Count();

            XWPFDocumentMapper mapper = new(paragraphMapperMock.Object, tableRowMapperMock.Object);
            mapper.MapTables(document, data);

            tableRowMapperMock.Verify(
                tm => tm.MapEnumerableToRow(It.IsAny<XWPFTableRow>(), It.IsAny<List<Dictionary<string, object>>>()),
                    Times.Exactly(0)
            );
        }

        [Fact]
        public void ItShould_MapNoEnumerables_IfMappingListIsNull()
        {
            Mock<IXWPFParagraphMapper> paragraphMapperMock = new();

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
                .Returns<XWPFTableRow, List<Dictionary<string, object>>>((r, d) => r); //GetMappingList
            tableRowMapperMock
                .Setup(t => t.GetMappingList(
                    It.IsAny<XWPFTableRow>(),
                    It.IsAny<KeyValuePair<string, IEnumerable<object>>>()
                ))
                .Returns<XWPFTableRow, KeyValuePair<string, IEnumerable<object>>>((r, d) => null!);

            List<object> theList = new() {
                new { A = "Chirp", B = "Meow", C = "Woof", D = "Moo", E = "Bonjour", F = "Quack"},
                new { A = "Monday", B = "Tuesday", C = "Wednesday", D = "Thursday", E = "Friday", F = "Saturday"},
                new { A = "Red", B = "Green", C = "Blue", D = "Yellow", E = "Magenta", F = "Cyan"}
            };
            Dictionary<string, object> data = new()
            {
                { "{{List}}", theList }
            };

            string template = @"TestDocuments/test3.docx";
            using FileStream fileStream = File.OpenRead(template);
            XWPFDocument document = new(fileStream);

            int rowCount = document.Tables.SelectMany(t => t.Rows).Select(r => r.GetTableCells()).Count();

            XWPFDocumentMapper mapper = new(paragraphMapperMock.Object, tableRowMapperMock.Object);
            mapper.MapTables(document, data);

            tableRowMapperMock.Verify(
                tm => tm.MapEnumerableToRow(It.IsAny<XWPFTableRow>(), It.IsAny<List<Dictionary<string, object>>>()),
                    Times.Exactly(0)
            );
        }
    }
}
