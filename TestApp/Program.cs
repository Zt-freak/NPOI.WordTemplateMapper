using NPOI.WordMapper.Extensions;
using NPOI.XWPF.UserModel;

// Assemble the data
List<object> theList = new() {
    new { A = "Chirp", B = "Meow", C = "Woof", D = "Moo", E = "Bonjour", F = "Quack"},
    new { A = "Monday", B = "Tuesday", C = "Wednesday", D = "Thursday", E = "Friday", F = "Saturday"},
    new { A = "Red", B = "Green", C = "Blue", D = "Yellow", E = "Magenta", F = "Cyan"},
    new { A = "Belgium", B = "Denmark", C = "Spain", D = "Austria", E = "Monaco", F = "Angola"},
    new { A = PokeType.Water, B = PokeType.Grass, C = PokeType.Fire, D = PokeType.Ghost, E = PokeType.Electric, F = PokeType.Fighting},
};
Dictionary<string, object> data = new()
{
    { "{{List}}", theList },
    { "{{NeverGivingYouUp}}", "Rick Astley" },
    { "{{ListTotal}}", 42069 },
    { "{{Pokemon}}", new { Name = "Pikachu" } },
    { "{{Phone}}", new { Brand = new { Name = "Huawei"} } },
    { "{{Actions}}", new string[] { "give you up", "let you down", "run around and desert you" } }
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