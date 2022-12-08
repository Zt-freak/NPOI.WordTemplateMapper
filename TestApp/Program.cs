using NPOI.WordMapper.Extensions;
using NPOI.XWPF.UserModel;

// Assemble the data
List<object> theList = new() {
    new { A = "chirp", B = "meow", C = "woof", D = "moo", E = "baah", F = "quack"},
    new { A = "monday", B = "tuesday", C = "wednesday", D = "thursday", E = "friday", F = "saturday"},
    new { A = "red", B = "green", C = "blue", D = "yellow", E = "magenta", F = "cyan"},
    new { A = "belgium", B = "denmark", C = "spain", D = "austria", E = "monaco", F = "angola"},
};
Dictionary<string, object> data = new()
{
    { "{{List}}", theList },
    { "{{SomeValue}}", "Rick Astley" },
    { "{{ListTotal}}", 42069 },
    { "{{Pokemon}}", new { Name = "Pikachu", PokeType = PokeType.Electric } }
};

// Get the Document
string template = @"A.docx";
using FileStream fileStream = File.OpenRead(template);
XWPFDocument document = new(fileStream);

// Map document
document.MapDocument(data);

// Create new document
string generateFile = @"output.docx";
using FileStream outputStream = File.Create(generateFile);
document.Write(outputStream);