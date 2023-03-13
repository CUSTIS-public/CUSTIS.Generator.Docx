using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json.Linq;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ApprovalTests;
using ApprovalTests.Namers;
using ApprovalTests.Reporters;
using CUSTIS.Generator.Docx;
using DocumentFormat.OpenXml.Packaging;

namespace CUSTIS.Generator.Docx.Tests;

[TestClass]
[UseReporter(typeof(DiffReporter))]
public class WordDocumentProcessorTest
{
    [TestMethod]
    public async Task PopulateSimpleDocument()
    {
        // Arrange
        var resultFile = "simple.filled.docx";
        File.Delete(resultFile);

        var input = await File.ReadAllTextAsync(Path.Combine(@"Samples", "simple.input.json"));

        var docProcessor = new WordDocumentProcessor(NullLogger<WordDocumentProcessor>.Instance);

        // Act
        using var filled = await docProcessor.PopulateDocumentTemplate(Path.Combine(@"Samples", "simple.template.docx"), input);
        await using (var resultFileStream = new FileStream(resultFile, FileMode.OpenOrCreate, FileAccess.Write))
        {
            await filled.CopyToAsync(resultFileStream);
        }

        // Assert
        await using var resultStream = new FileStream(resultFile, FileMode.Open, FileAccess.Read);
        using var doc = WordprocessingDocument.Open(resultStream, false);

        NamerFactory.AdditionalInformation = "doc";
        Approvals.VerifyXml(doc.MainDocumentPart?.Document.OuterXml);

        NamerFactory.AdditionalInformation = "num";
        Approvals.VerifyXml(doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering?.OuterXml);
    }

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
            docProcessor.PopulateDocumentTemplate(fileStream, input);
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
            docProcessor.PopulateDocumentTemplate(fileStream, input);
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
            docProcessor.PopulateDocumentTemplate(fileStream, input);
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
    [DataRow(true)]
    [DataRow(false)]
    public void PopulateComplexDocumentWithErrors(bool showErrorsInDocument)
    {
        // Arrange
        var additionalInformation = showErrorsInDocument ? "showErrors" : "hideErrors";
        var resultFile = $"ComplexDocument.{additionalInformation}.filled.docx";
        File.Delete(resultFile);
        File.Copy(Path.Combine(@"Samples", "ComplexDocument.template.docx"), resultFile);
        var input = new JObject(); // пустой объект

        var docProcessor = new WordDocumentProcessor(NullLogger<WordDocumentProcessor>.Instance);

        // Act
        using (var fileStream = new FileStream(resultFile, FileMode.Open, FileAccess.ReadWrite))
        {
            docProcessor.PopulateDocumentTemplate(fileStream, input, showErrorsInDocument);
        }

        // Assert
        using var resultStream = new FileStream(resultFile, FileMode.Open, FileAccess.ReadWrite);
        using var doc = WordprocessingDocument.Open(resultStream, false);

        NamerFactory.AdditionalInformation = additionalInformation;
        Approvals.VerifyXml(doc.MainDocumentPart?.Document.OuterXml);
    }

    [DataTestMethod]
    [DataRow(true)]
    [DataRow(false)]
    public void PopulateDocumentWithErrors(bool showErrorsInDocument)
    {
        // Arrange
        var additionalInformation = showErrorsInDocument ? "showErrors" : "hideErrors";
        var resultFile = $"errors.{additionalInformation}.filled.docx";

        File.Delete(resultFile);
        File.Copy(Path.Combine(@"Samples", "errors.template.docx"), resultFile);
        var input = JObject.Parse(File.ReadAllText(Path.Combine(@"Samples", $"errors.input.json")));
        var docProcessor = new WordDocumentProcessor(NullLogger<WordDocumentProcessor>.Instance);

        // Act
        using (var fileStream = new FileStream(resultFile, FileMode.Open, FileAccess.ReadWrite))
        {
            docProcessor.PopulateDocumentTemplate(fileStream, input, showErrorsInDocument);
        }

        // Assert
        using var resultStream = new FileStream(resultFile, FileMode.Open, FileAccess.ReadWrite);
        using var doc = WordprocessingDocument.Open(resultStream, false);

        NamerFactory.AdditionalInformation = additionalInformation;
        Approvals.VerifyXml(doc.MainDocumentPart?.Document.OuterXml);
    }

    [TestMethod]
    public async Task TestConditionalVisibility()
    {
        // Arrange
        var resultFile = "conditional.filled.docx";
        File.Delete(resultFile);

        var input = await File.ReadAllTextAsync(Path.Combine(@"Samples", "conditional.input.json"));

        var docProcessor = new WordDocumentProcessor(NullLogger<WordDocumentProcessor>.Instance);

        // Act
        using var filled = await docProcessor.PopulateDocumentTemplate(Path.Combine(@"Samples", "conditional.template.docx"), input, true);
        await using (var resultFileStream = new FileStream(resultFile, FileMode.OpenOrCreate, FileAccess.Write))
        {
            await filled.CopyToAsync(resultFileStream);
        }

        // Assert
        await using var resultStream = new FileStream(resultFile, FileMode.Open, FileAccess.Read);
        using var doc = WordprocessingDocument.Open(resultStream, false);

        NamerFactory.AdditionalInformation = "doc";
        Approvals.VerifyXml(doc.MainDocumentPart?.Document.OuterXml);

        NamerFactory.AdditionalInformation = "num";
        Approvals.VerifyXml(doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering?.OuterXml);
    }

    [TestMethod]
    [Ignore("need to repair Approvals.VerifyXml")]
    public async Task LineBreakTest()
    {
        // Arrange
        var resultFile = "line-break.filled.docx";
        File.Delete(resultFile);

        var input = await File.ReadAllTextAsync(Path.Combine(@"Samples", "line-break.input.json"));

        var docProcessor = new WordDocumentProcessor(NullLogger<WordDocumentProcessor>.Instance,
            new ProcessorOptions { ReplaceLineBreakWithTag = true });

        // Act
        using var filled = await docProcessor.PopulateDocumentTemplate(Path.Combine(@"Samples", "line-break.template.docx"), input);
        await using (var resultFileStream = new FileStream(resultFile, FileMode.OpenOrCreate, FileAccess.Write))
        {
            await filled.CopyToAsync(resultFileStream);
        }

        // Assert
        await using var resultStream = new FileStream(resultFile, FileMode.Open, FileAccess.Read);
        using var doc = WordprocessingDocument.Open(resultStream, false);

        NamerFactory.AdditionalInformation = "doc";
        Approvals.VerifyXml(doc.MainDocumentPart?.Document.OuterXml);

        NamerFactory.AdditionalInformation = "num";
        Approvals.VerifyXml(doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering?.OuterXml);
    }

    [TestMethod]
    public async Task StreamWithJsonTest()
    {
        // Arrange
        var resultFile = "StreamWithJsonTest.filled.docx";
        File.Delete(resultFile);

        var input = new
        {
            textInRun = "Line1\r\nLine2\rLine3\nLine4",
            textInRunInParagraph = "Line1\r\nLine2\rLine3\nLine4",
            textInRunInParagraphInCell = "Line1\r\nLine2\rLine3\nLine4",
            textInRunAllowMulti = "Line1\r\nLine2\rLine3\nLine4",
            textInRunWithPlaceholderText = "Line1\r\nLine2\rLine3\nLine4"
        };

        var stream = new FileStream(Path.Combine(@"Samples", "line-break.template.docx"), FileMode.Open, FileAccess.Read);

        var docProcessor = new WordDocumentProcessor(NullLogger<WordDocumentProcessor>.Instance,
            new ProcessorOptions { ReplaceLineBreakWithTag = true });

        // Act - Assert
        using var filled = await docProcessor.PopulateDocumentTemplate(stream, input);
        await using (var resultFileStream = new FileStream(resultFile, FileMode.OpenOrCreate, FileAccess.Write))
        {
            await filled.CopyToAsync(resultFileStream);
        }
    }
}