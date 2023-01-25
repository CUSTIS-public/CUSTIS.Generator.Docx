using System.IO;
using CUSTIS.Generator.Docx;
using Newtonsoft.Json.Linq;
using Microsoft.Extensions.Logging.Abstractions;

Console.WriteLine("Welcome to CUSTIS.Generator.DocX!");

File.Copy("SimpleTemplate.docx", "SimpleTemplate.filled.docx");

using var fileStream = new FileStream("SimpleTemplate.filled.docx", FileMode.Open, FileAccess.ReadWrite);
var input = new JObject
{
    ["textInRun"] = "Text in Run",
    ["textInRunInParagraph"] = "Text in Run in Paragraph",
    ["textInRunInParagraphInCell"] = "Text in Run in Paragraph in Cell",
    ["textInRunAllowMulti"] = "Text in Run Allow Multi Line 1\r\nLine 2.",
    ["textInRunWithPlaceholderText"] = "Text in Run with Placeholder Text",
};

var docProcessor = new WordDocumentProcessor(NullLogger<WordDocumentProcessor>.Instance);
docProcessor.PopulateDocumentTemplate(fileStream, input);

Console.WriteLine("Template successfully filled and stored as SimpleDocument.filled.docx");