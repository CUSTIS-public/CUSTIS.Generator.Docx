using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json.Linq;
using System.IO;
using System.Linq;
using ApprovalTests;
using ApprovalTests.Namers;
using ApprovalTests.Reporters;
using CUSTIS.DocumentService;
using DocumentFormat.OpenXml.Packaging;

namespace Novo.DocumentService.Tests;

[TestClass]
[UseReporter(typeof(DiffReporter))]
public class WordDocumentProcessorTest
{
    [TestMethod]
    public void PopulateComplexDocument()
    {
        // Arrange
        var resultFile = "ComplexDocument.filled.docx";
        File.Delete(resultFile);
        File.Copy(Path.Combine(@"Samples", "ComplexDocument.template.docx"), resultFile);
        var input = JObject.Parse(File.ReadAllText(Path.Combine(@"Samples", $"ComplexDocument.input.json")));

        var docProcessor = new WordDocumentProcessor(NullLogger<WordDocumentProcessor>.Instance);

        // Act
        using (var fileStream = new FileStream(resultFile, FileMode.Open, FileAccess.ReadWrite))
        {
            docProcessor.PopulateDocumentTemplate(input, fileStream);
        }

        // Assert
        using var resultStream = new FileStream(resultFile, FileMode.Open, FileAccess.ReadWrite);
        using var doc = WordprocessingDocument.Open(resultStream, false);

        NamerFactory.AdditionalInformation = "doc";
        Approvals.VerifyXml(doc.MainDocumentPart?.Document.OuterXml);

        NamerFactory.AdditionalInformation = "num";
        Approvals.VerifyXml(doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering?.OuterXml);
    }

    [DataTestMethod]
    [DataRow("simple_text")]
    [DataRow("paragraphs")]
    [DataRow("br")]
    [DataRow("ul")]
    [DataRow("nested_ul")]
    [DataRow("ol")]
    [DataRow("nested_ol")]
    [DataRow("invalid_tags")]
    public void PopulateDocumentWithHtml(string caseName)
    {
        // Arrange
        var resultFile = $"html.{caseName}.filled.docx";
        File.Delete(resultFile);
        File.Copy(Path.Combine(@"Samples", "html.template.docx"), resultFile);
        var input = JObject.Parse(File.ReadAllText(Path.Combine(@"Samples", $"html.InputParameters.{caseName}.json")));

        var docProcessor = new WordDocumentProcessor(NullLogger<WordDocumentProcessor>.Instance);

        // Act
        using (var fileStream = new FileStream(resultFile, FileMode.Open, FileAccess.ReadWrite))
        {
            docProcessor.PopulateDocumentTemplate(input, fileStream);
        }

        // Assert
        using var resultStream = new FileStream(resultFile, FileMode.Open, FileAccess.ReadWrite);
        using var doc = WordprocessingDocument.Open(resultStream, false);

        NamerFactory.AdditionalInformation = $"{caseName}.doc";
        Approvals.VerifyXml(doc.MainDocumentPart?.Document.OuterXml);

        NamerFactory.AdditionalInformation = $"{caseName}.num";
        Approvals.VerifyXml(doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering?.OuterXml);
    }

    [TestMethod]
    public void PopulateHtmlArrayDocument()
    {
        // Arrange
        var resultFile = "html-array.filled.docx";
        File.Delete(resultFile);
        File.Copy(Path.Combine(@"Samples", "html-array.template.docx"), resultFile);
        var items = File.ReadAllLines(Path.Combine(@"Samples", $"html-array.input.txt"));
        var input = new JObject
        {
            ["items"] = new JArray(items.Select((item, i) => new JObject
            {
                ["head"] = $"Item {i}",
                ["html"] = item
            }))
        };

        var docProcessor = new WordDocumentProcessor(NullLogger<WordDocumentProcessor>.Instance);

        // Act
        using (var fileStream = new FileStream(resultFile, FileMode.Open, FileAccess.ReadWrite))
        {
            docProcessor.PopulateDocumentTemplate(input, fileStream);
        }

        // Assert
        using var resultStream = new FileStream(resultFile, FileMode.Open, FileAccess.ReadWrite);
        using var doc = WordprocessingDocument.Open(resultStream, false);

        NamerFactory.AdditionalInformation = "doc";
        Approvals.VerifyXml(doc.MainDocumentPart?.Document.OuterXml);

        NamerFactory.AdditionalInformation = "num";
        Approvals.VerifyXml(doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering?.OuterXml);
    }
}