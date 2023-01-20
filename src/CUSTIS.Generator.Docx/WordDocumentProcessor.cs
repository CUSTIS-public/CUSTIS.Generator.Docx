using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;

namespace CUSTIS.DocumentService;

public interface IDocumentProcessor
{
    DocumentProcessingResult PopulateDocumentTemplate(JObject input);
    void PopulateDocumentTemplate(JObject parameters, Stream stream);
}

public class WordDocumentProcessor : IDocumentProcessor
{
    const string documentRelationshipNs =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

    const string stylesRelationshipNs =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

    const string wordml2006Ns =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    const string wordml2012Ns =
        "http://schemas.microsoft.com/office/word/2012/wordml";


    private readonly ILogger<WordDocumentProcessor> _logger;
    private static readonly XNamespace W = (XNamespace)wordml2006Ns;

    public WordDocumentProcessor(ILogger<WordDocumentProcessor> log) => _logger = log;

    public DocumentProcessingResult PopulateDocumentTemplate(JObject input)
    {
        if (!input.TryGetValue<JObject>("parameters", out var parameters))
        {
            return new DocumentProcessingResult(
                false,
                new Error(
                    "The 'parameters' property was not found in the input. You need to provide an input in the" +
                    " format {\"parameters\":{ ... }, \"file\":\"base64 encoded docx file\""));
        }

        if (parameters == null)
        {
            return new DocumentProcessingResult(false, new Error(
                "The `parameters' property has no values. You need to provide an input in the format" +
                " {\"parameters\":{ ... }, \"file\":\"base64 encoded docx file\""));
        }

        if (!input.TryGetValue("file", out var inputfile) || inputfile.Type != JTokenType.String)
        {
            return new DocumentProcessingResult(
                false,
                new Error(
                    "The 'file' token was not found in the input. You need to provide an input in the format" +
                    " {\"parameters\":{ ... }, \"file\":\"base64 encoded docx file\""));
        }

        var base64 = inputfile.ToString();
        var buffer = new byte[(base64.Length * 3 + 3) / 4 -
                              (base64.Length > 0 && base64[^1] == '='
                                  ? base64.Length > 1 && base64[^2] == '=' ? 2 : 1
                                  : 0)];
        if (!Convert.TryFromBase64String(base64, buffer, out int writtenBytes))
        {
            return new DocumentProcessingResult(
                success: false,
                result: new ParsingError(
                    message:
                    "Document is not in base64 format. Make sure you are " +
                    "sending a valid base64 encoded docx file.",
                    bytesParsed: writtenBytes));
        }

        using var stream = new MemoryStream();
        stream.Write(buffer, 0, buffer.Length);
        try
        {
            PopulateDocumentTemplate(parameters, stream);
        }
        catch (Exception ex)
        {
            _logger.LogCritical($"Unable to populate the document due to an error. {ex}");
            throw;
        }

        return new DocumentProcessingResult(
            true,
            new JObject(new JProperty("file", Convert.ToBase64String(stream.ToArray()))));
    }

    public void PopulateDocumentTemplate(JObject parameters, Stream stream)
    {
        using var doc = WordprocessingDocument.Open(stream, true);
        var mainPart = doc.MainDocumentPart;
        if (mainPart == null)
        {
            throw new InvalidOperationException(
                "Invalid document format. The document is lacking its main part.");
        }

        var topLevelSdtElements
            = mainPart.Document.Descendants<SdtElement>().Where(e => !e.Ancestors<SdtElement>().Any());
        // Note: An <sdt> can be SdtBlock, SdtCell, SdtRow, SdtRun, SdtRunRub, and they are inherit
        // from SdtElement. That's why we used SdtElement above.
        foreach (var sdtElement in topLevelSdtElements)
        {
            PopulateSdtElement(parameters, sdtElement);
        }

        doc.Save();
    }

    private void PopulateSdtElement(JObject parameters, SdtElement sdtElement)
    {
        if (parameters == null)
        {
            throw new ArgumentNullException(nameof(parameters));
        }

        if (sdtElement == null)
        {
            throw new ArgumentNullException(nameof(sdtElement));
        }

        var tag = sdtElement.SdtProperties?.Descendants<Tag>().FirstOrDefault();
        if (tag == null || string.IsNullOrWhiteSpace(tag.Val))
        {
            _logger.LogWarning(
                "Placeholder found without tag. Placeholders without tag are not supported.");
            return;
        }

        if (string.IsNullOrWhiteSpace(tag.Val))
        {
            _logger.LogWarning(
                "Placeholder found without an empty tag. Placeholders without tag are not supported.");
            return;
        }

        _logger.LogDebug($"Processing tag: '{tag.Val}'");
        if (!parameters.TryGetValue(tag.Val!, out var token))
        {
            try
            {
                token = parameters.SelectToken(tag.Val!);
                if (token == null)
                {
                    _logger.LogInformation($"No parameter matched placeholder '{tag.Val}'.");
                    return;
                }
            }
            catch (JsonException e) when (e.Message.Contains("Path returned multiple tokens"))
            {
                token = new JArray(parameters.SelectTokens(tag.Val!));
            }
        }

        try
        {
            if (sdtElement.SdtProperties?.Descendants<SdtContentText>().Any() ?? false)
            {
                // if <w:text> is present, then this is plain text control
                ProcessText(sdtElement, tag, token);
            }
            else if (sdtElement.SdtProperties!.Descendants<SdtRepeatedSection>().Any())
            {
                // if <w15:repeatingSection> is present, then this is repeating section
                ProcessRepeatedSection(sdtElement, tag, token);
            }
            else
            {
                // we'll suggest this is rich text control, but this could also be any other control
                // the problem is Word puts no special property for rich text
                ProcessRichText(sdtElement, tag, token); 
            }
        }
        catch (Exception e)
        {
            _logger.LogError(e, "An error '{Message}' while processing '{Tag}'", e.Message, tag.Val);
        }
    }

    private void ProcessRepeatedSection(SdtElement sdtElement, Tag tag, JToken token)
    {
        // Repeating Section Content Control
        if (token is not JArray tokens)
        {
            _logger.LogWarning(
                $"The value of '{tag.Val}' parameter is not an array. Parameter mapped " +
                $"to a repeating section can only be an array.");
            return;
        }

        var content = sdtElement.Elements().FirstOrDefault(e => e.XName == W + "sdtContent");
        if (content == null || content.FirstChild == null)
        {
            _logger.LogCritical(
                $"Encountered repeating with no content area for repeating item. It will be skipped.");
        }

        if (content!.FirstChild is not SdtElement firstRepeatingItem)
        {
            _logger.LogWarning($"Encountered repeating with wrong element instead of repeating item." +
                               $" It will be skipped.");
            return;
        }

        // Remove all exisitng repeating items (in the future we can allow the user to turn it on/off).
        content.RemoveAllChildren();

        //foreach (var id in repeatingItemClone.Descendants<SdtId>())
        foreach (var id in firstRepeatingItem.Descendants<SdtId>())
        {
            id.Remove();
        }

        foreach (JObject tokenChild in tokens)
        {
            // Find the first repeating section item, and clone it.
            var repeatingItemClone = firstRepeatingItem.CloneNode(true);
            var repeatingItemChildSdtElements =
                repeatingItemClone.Descendants<SdtElement>().Where(
                    e => !e.Ancestors<SdtElement>().Except(new[] { repeatingItemClone }).Any()).ToArray();
            content.AppendChild(repeatingItemClone);
            foreach (var sdt in repeatingItemChildSdtElements)
            {
                PopulateSdtElement(tokenChild, sdt);
            }

            //lastRepeatingItem.InsertAfterSelf(repeatingItemClone);
            //// If the last repeating section item is empty remove it.
            //if (!lastRepeatingItem.Descendants<SdtElement>().Any(
            //    e => e.SdtProperties?.Elements<ShowingPlaceholder>().SingleOrDefault() == null))
            //{
            //    lastRepeatingItem.Remove();
            //}
            _logger.LogDebug("Placeholder row inserted successfully.");
        }
    }

    private void ProcessText(SdtElement sdtElement, Tag tag, JToken token)
    {
        _logger.LogDebug("{Tag}: {Token}", tag.Val, token);

        // Plain or Rich Text Content Control
        sdtElement.SdtProperties?.Elements<ShowingPlaceholder>().FirstOrDefault()?.Remove();
        // There are several possible types for <sdtContent> element (e.g. SdtContentBlock)
        // That's why we don't use a concrete type in the following line.
        var contentElement = FindContent(sdtElement);
        if (contentElement == null)
        {
            _logger.LogWarning("Placeholder doesn't have any content area.");
            return;
        }

        // There can be at most one paragraph.
        var paragraph = contentElement.Descendants<Paragraph>().SingleOrDefault();
        // There can be multiple runs, but we keep only the first.
        var firstRun = contentElement.Descendants<Run>()?.FirstOrDefault();
        var runs = contentElement.Elements<Run>()?.Skip(1);
        if (runs != null)
        {
            foreach (var run in runs)
            {
                run.Remove();
            }
        }

        // There can be multiple texts, but we keep only the first.
        var firstText = contentElement.Descendants<Text>()?.FirstOrDefault();
        var texts = contentElement.Elements<Run>()?.Skip(1);
        if (texts != null)
        {
            foreach (var text in texts)
            {
                text.Remove();
            }
        }

        var strToken = token.ToString();

        if (firstText == null)
        {
            if (firstRun != null)
            {
                firstRun.AddChild(new Text(strToken));
            }
            else if (paragraph != null)
            {
                paragraph.AddChild(new Run(new Text(strToken)));
            }
            else
            {
                _logger.LogWarning($"Place holder '{tag.Val}' does not have a correct structure.");
                return;
            }
        }
        else
        {
            firstText.Text = strToken;
            firstText.Parent!.Descendants<RunStyle>().FirstOrDefault(s => s.Val == "PlaceholderText")?.Remove();
        }

        paragraph?.Descendants<RunStyle>().FirstOrDefault(s => s.Val == "PlaceholderText")?.Remove();
        firstRun?.Descendants<RunStyle>().FirstOrDefault(s => s.Val == "PlaceholderText")?.Remove();
    }

    private static OpenXmlElement? FindContent(SdtElement sdtElement)
    {
        return sdtElement.Descendants().SingleOrDefault(e => e.XName == W + "sdtContent");
    }

    private void ProcessRichText(SdtElement sdtElement, Tag tag, JToken content)
    {
        _logger.LogDebug("RichText: {Tag}: {Content}", tag.Val, content);

        if (sdtElement.Descendants<SdtContentText>().Any())
        {
            _logger.LogWarning("HTML cannot be written to PlainText control. Use Rich Text Control instead");
            return;
        }
        
        var contentElement = FindContent(sdtElement);
        if (contentElement == null)
        {
            _logger.LogWarning("Placeholder doesn't have any content area.");
            return;
        }

        var mainDocumentPart = contentElement.Ancestors<Document>().First().MainDocumentPart ?? throw new NullReferenceException($"Failed to obtain MainDocumentPart from {contentElement}");
        var docx = content.ToString().ConvertToDocx(mainDocumentPart);

        if (docx.AbstractNums.Any() || docx.NumberingInstances.Any())
        {
            var numberingPart = mainDocumentPart.NumberingDefinitionsPart;
            if (numberingPart == null)
            {
                numberingPart = mainDocumentPart.AddNewPart<NumberingDefinitionsPart>("NumberingDefinitionsPart001");
                new Numbering().Save(numberingPart);
            }

            Insert(numberingPart, docx.AbstractNums);
            Insert(numberingPart, docx.NumberingInstances);
        }

        var cell = contentElement.Descendants<TableCell>().FirstOrDefault();
        var pasteHere = cell ?? contentElement;
        pasteHere.RemoveAllChildren();
        pasteHere.Append(docx.Paragraphs);
    }

    private static void Insert<T>(NumberingDefinitionsPart numberingPart, IList<T> elements) 
        where T : OpenXmlElement
    {
        var lastElement = numberingPart.Numbering.Elements<T>().LastOrDefault();
        if (lastElement != null)
        {
            foreach (var element in elements)
            {
                numberingPart.Numbering.InsertAfter(element, lastElement);
                lastElement = element;
            }
        }
        else
        {
            numberingPart.Numbering.Append(elements);
        }
    }
}