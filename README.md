# NPOI.WordTemplateMapper

This is a Word document mapper which can map a disctionary of objects to a document.

## usage

Example usage

```csharp
using NPOI.WordTemplateMapper.XWPF;
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
    { "{{CoolSinger}}", "Rick Astley" },
    { "{{ListTotal}}", 42069 },
    { "{{Pokemon}}", new { Name = "Pikachu", PokeType = PokeType.Electric } },
    { "{{Phone}}", new { Brand = new { Name = "Huawei"}, Model = "Mate 50 Pro" } },
    { "{{Actions}}", new string[] { "give you up", "let you down", "run around and desert you" } },
    { "{{Wow}}", new Array[] { new object[] { new { Yay = "waaaa" } } } }
};

// Get the Document
string template = @"A.docx";
using FileStream fileStream = File.OpenRead(template);
XWPFDocument document = new(fileStream);

// Map document
XWPFDocumentMapper documentMapper = new();

documentMapper.MapDocument(document, data);

// Create new document
string generateFile = @"output.docx";
using FileStream outputStream = File.Create(generateFile);
document.Write(outputStream);
```

### Mapping

Mapping keys must start and end with non alphanumeric characters, preferebly `{{two accolades like these}}`.

### Mapping table rows

Add your key to the title of the table. This can be accessed under "Table Properties" in Microsoft Word. This key will be cleaned up during mapping, so won't need to worry about accessibility.