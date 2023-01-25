[![NuGet version (CUSTIS.Generator.Docx)](https://img.shields.io/nuget/v/CUSTIS.Generator.Docx.svg?style=flat-square)](https://www.nuget.org/packages/CUSTIS.Generator.Docx)

# CUSTIS.Generator.Docx 
Generates docx document from docx template using data in JSON-format. This is a port from [novodocx](https://github.com/rezanid/novodocx) with some improvements.

# Getting started

1. Install `CUSTIS.Generator.Docx` [package](https://www.nuget.org/packages/CUSTIS.Generator.Docx)
2. Prepare docx-template (see [instructions](#how-to-create-template) below)
3. Populate template with your data:

```csharp
var jsonData = "{'someData': 'hello'}"; // data to populate template with
var templateFile = "template.docx"; // file with template 

// populate file with data
var docProcessor = new WordDocumentProcessor(NullLogger<WordDocumentProcessor>.Instance);
using var filled = await docProcessor.PopulateDocumentTemplate(templateFile, input);

// save file
await using var resultFileStream = new FileStream("destination.docx", FileMode.OpenOrCreate, FileAccess.Write);
await filled.CopyToAsync(resultFileStream);
```

# How to create template

Open Developer Tab in MS Word and enter Design Mode ([instructions](https://www.namicsoft.com/doc/enter-design-mode-microsoft-word/)).

## Simple text

![Plain text](.media/plain_text.png)

1. Use `Plain Text Content Control` if you need to render text without line breaks
2. Open `Properties`
3. Fill `Tag` with the field name from JSON

Data sample:
```csharp
var jsonData = @"{'name': 'CUSTIS.Generator.Docx'}";
```

Output: `Hello, CUSTIS.Generator.Docx!`

## HTML

![HTML](.media/html.png)

1. Use `Rich Text Content Control` if you need to render HTML
2. Open `Properties`
3. Fill `Tag` with the field name from JSON

Limitations: only `<p>`, `<br>`, `<ol>`, `<ul>`, `<li>` tags are processed. All other tags are skipped.

Data sample:
```csharp
var jsonData = @"{'formatted': 'Hello,<br/>CUSTIS.<red>Generator</red>.Docx!'}";
```

Output: 
```
Hello,
CUSTIS.Generator.Docx!
```

## Repeated data (arrays)

![Repeated](.media/repeated.png)

1. Use `Repeated Section Control` if you need to render array
2. Open `Properties`
3. Fill `Tag` with the name of array from JSON
4. Use either `Plain` or `Rich Text Content Controls` inside `Repeated Section Control` to render text 

Arrays can be rendered as paragraphs, lists or even tables. You can even render nested arrays!

Data sample:
```csharp
var jsonData = @"{'sports': [{'name': 'swimming'}, {'name': 'football'}]}";
```

Output:
```
Sports:
1. Name: swimming
2. Name: football
```

# Advanced features

[Json Path expressions](https://goessner.net/articles/JsonPath) can be used to query data. `CUSTIS.Generator.Docx` uses `Json.NET` (from Newtonsoft). You can get some samples of JSON path [here](https://www.newtonsoft.com/json/help/html/QueryJsonSelectToken.htm). 

# Links

* [Some info about creating Word templates](https://support.clio.com/hc/en-us/articles/203359420-Tutorial-Creating-a-Protected-Fillable-Template-in-Microsoft-Word-for-Use-in-Document-Automation)