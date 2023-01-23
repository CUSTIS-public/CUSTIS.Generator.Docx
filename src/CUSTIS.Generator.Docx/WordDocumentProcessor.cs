using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Color = DocumentFormat.OpenXml.Wordprocessing.Color;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;

namespace CUSTIS.Generator.Docx;

public class WordDocumentProcessor : IDocumentProcessor
{
    private const string Wordml2006Ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string Red = "C00000";

    private readonly ILogger<WordDocumentProcessor> _logger;
    private static readonly XNamespace W = Wordml2006Ns;

    public WordDocumentProcessor(ILogger<WordDocumentProcessor> log) => _logger = log;

    public void PopulateDocumentTemplate(Stream stream, JObject parameters, bool showErrorsInDocument = false)
    {
        using var doc = WordprocessingDocument.Open(stream, true);
        var mainPart = doc.MainDocumentPart;
        if (mainPart == null)
        {
            throw new InvalidOperationException(
                "Invalid document format. The document is lacking its main part.");
        }

        var errors = new ErrorsCollector();

        var topLevelSdtElements
            = mainPart.Document.Descendants<SdtElement>().Where(e => !e.Ancestors<SdtElement>().Any());
        // Note: An <sdt> can be SdtBlock, SdtCell, SdtRow, SdtRun, SdtRunRub, and they are inherit
        // from SdtElement. That's why we used SdtElement above.
        foreach (var sdtElement in topLevelSdtElements)
        {
            PopulateSdtElement(parameters, sdtElement, errors);
        }

        if (showErrorsInDocument && errors.Errors.Any())
        {
            var bold = new Bold();
            var color = new Color() { Val = Red };
            var size = new FontSize() { Val = "36" };
            var text = new Text("Some errors occured while generating document");
            var runPro = new RunProperties(bold, color, size, text);
            OpenXmlElement lastError = new Paragraph(new Run(runPro));
            mainPart.Document.PrependChild(lastError);
            foreach (var error in errors.Errors)
            {
                var newChild = new Paragraph(new Run(new Text(error.Message)));
                lastError.InsertAfterSelf(newChild);
                lastError = newChild;

                if (error.Element is SdtElement sdtElement)
                {
                    foreach (var runProps in sdtElement.Descendants<RunProperties>())
                    {
                        runProps.PrependChild(new Color() { Val = Red });
                        runProps.PrependChild(new Bold());
                    }
                }
            }

            var space = new Paragraph(new Run(new Text()));
            lastError.InsertAfterSelf(space);
        }

        doc.Save();
    }

    private void PopulateSdtElement(JObject parameters, SdtElement sdtElement, ErrorsCollector errorsCollector)
    {
        if (parameters == null)
        {
            throw new ArgumentNullException(nameof(parameters));
        }

        if (sdtElement == null)
        {
            throw new ArgumentNullException(nameof(sdtElement));
        }

        var sdtProperties = sdtElement.SdtProperties;
        var tag = sdtProperties?.Descendants<Tag>().FirstOrDefault();
        if (sdtProperties == null || tag == null || string.IsNullOrWhiteSpace(tag.Val))
        {
            errorsCollector.AddError("Placeholder found without tag. Placeholders without tag are not supported",
                sdtElement);
            _logger.LogWarning("Placeholder found without tag. Placeholders without tag are not supported");
            return;
        }

        if (string.IsNullOrWhiteSpace(tag.Val))
        {
            errorsCollector.AddError("Placeholder found with an empty tag. Placeholders without tag are not supported",
                sdtElement);
            _logger.LogWarning("Placeholder found with an empty tag. Placeholders without tag are not supported");
            return;
        }

        _logger.LogDebug("Processing tag: '{Tag}'", tag.Val);
        if (!parameters.TryGetValue(tag.Val!, out var token))
        {
            try
            {
                token = parameters.SelectToken(tag.Val!);
                if (token == null)
                {
                    errorsCollector.AddError($"No data matched placeholder '{tag.Val}'", sdtElement);
                    _logger.LogInformation("No data matched placeholder '{Tag}'", tag.Val);
                    return;
                }
            }
            catch (JsonException e) when (e.Message.Contains("Path returned multiple tokens"))
            {
                token = new JArray(parameters.SelectTokens(tag.Val!));
            }
        }

        var unsupportedTypes = new[]
        {
            typeof(DocumentFormat.OpenXml.Wordprocessing.DataBinding),
            typeof(DocumentFormat.OpenXml.Office2013.Word.DataBinding),
            typeof(DocumentFormat.OpenXml.Wordprocessing.SdtContentEquation),
            typeof(DocumentFormat.OpenXml.Wordprocessing.SdtContentPicture),
            typeof(DocumentFormat.OpenXml.Wordprocessing.SdtContentCitation),
            typeof(DocumentFormat.OpenXml.Wordprocessing.SdtContentGroup),
            typeof(DocumentFormat.OpenXml.Wordprocessing.SdtContentBibliography),
            typeof(DocumentFormat.OpenXml.Office2010.Word.EntityPickerEmpty),
            typeof(DocumentFormat.OpenXml.Office2013.Word.SdtRepeatedSectionItem),
            typeof(DocumentFormat.OpenXml.Office2013.Word.WebExtensionLinked),
            typeof(DocumentFormat.OpenXml.Office2013.Word.WebExtensionCreated),
            typeof(DocumentFormat.OpenXml.Wordprocessing.SdtContentComboBox),
            typeof(DocumentFormat.OpenXml.Wordprocessing.SdtContentDate),
            typeof(DocumentFormat.OpenXml.Wordprocessing.SdtContentDocPartObject),
            typeof(DocumentFormat.OpenXml.Wordprocessing.SdtContentDocPartList),
            typeof(DocumentFormat.OpenXml.Wordprocessing.SdtContentDropDownList),
            typeof(DocumentFormat.OpenXml.Office2010.Word.SdtContentCheckBox),
            typeof(DocumentFormat.OpenXml.Office2013.Word.Appearance),
        };

        try
        {
            var unsupportedDescendants = sdtProperties.Descendants().Where(d => unsupportedTypes.Contains(d.GetType())).ToArray();
            if (unsupportedDescendants.Any())
            {
                var typeNames = unsupportedDescendants.Select(d => d.GetType().Name);
                _logger.LogWarning("Placeholder '{Tag}' has incorrect type '{Type}'. Only plain text, rich text and repeated section are supported", tag.Val, typeNames);
                errorsCollector.AddError($"Placeholder '{tag.Val}' has incorrect type '{string.Join(", ", typeNames)}'. Only plain text, rich text and repeated section are supported", sdtElement);
            }
            else if (sdtProperties.Descendants<SdtContentText>().Any())
            {
                // if <w:text> is present, then this is plain text control
                ProcessText(sdtElement, tag, token, errorsCollector);
            }
            else if (sdtProperties.Descendants<SdtRepeatedSection>().Any())
            {
                // if <w15:repeatingSection> is present, then this is repeating section
                ProcessRepeatedSection(sdtElement, tag, token, errorsCollector);
            }
            else
            {
                // sometimes Word puts no special property for rich text,
                // so we can't write sdtProperties.Descendants<SdtRichText>().Any()
                ProcessRichText(sdtElement, tag, token, errorsCollector);
            }
        }
        catch (Exception e)
        {
            errorsCollector.AddError($"An error '{e.Message}' while processing '{tag.Val}'", sdtElement);
            _logger.LogError(e, "An error '{Message}' while processing '{Tag}'", e.Message, tag.Val);
        }
    }

    private void ProcessRepeatedSection(SdtElement sdtElement, Tag tag, JToken token, ErrorsCollector errorsCollector)
    {
        // Repeating Section Content Control
        if (token is not JArray tokens)
        {
            errorsCollector.AddError(
                $"The value of '{tag.Val}' parameter is not an array. Parameter mapped to a repeating section can only be an array",
                sdtElement);
            _logger.LogWarning(
                "The value of '{Tag}' parameter is not an array. Parameter mapped to a repeating section can only be an array",
                tag.Val);
            return;
        }

        var content = sdtElement.Elements().FirstOrDefault(e => e.XName == W + "sdtContent");
        if (content?.FirstChild == null)
        {
            errorsCollector.AddError(
                $"Encountered repeating '{tag.Val}' with no content area for repeating item. It will be skipped",
                sdtElement);
            _logger.LogWarning(
                "Encountered repeating '{Tag}' with no content area for repeating item. It will be skipped", tag.Val);
        }

        if (content!.FirstChild is not SdtElement firstRepeatingItem)
        {
            errorsCollector.AddError(
                $"Encountered repeating '{tag.Val}' with wrong element instead of repeating item. It will be skipped",
                sdtElement);
            _logger.LogWarning(
                "Encountered repeating '{Tag}' with wrong element instead of repeating item. It will be skipped",
                tag.Val);
            return;
        }

        // Remove all exisitng repeating items (in the future we can allow the user to turn it on/off).
        content.RemoveAllChildren();

        //foreach (var id in repeatingItemClone.Descendants<SdtId>())
        foreach (var id in firstRepeatingItem.Descendants<SdtId>())
        {
            id.Remove();
        }

        foreach (var tokenChild in tokens)
        {
            if (tokenChild is not JObject jObject)
            {
                _logger.LogWarning("Element '{Token}' in '{Tag}' is not an object. It should be an object in '{{}}' braces", tokenChild, tag.Val);
                errorsCollector.AddError($"Element '{tokenChild}' in '{tag.Val}' is not an object. It should be an object in '{{}}' braces", sdtElement);
                continue;
            }
            // Find the first repeating section item, and clone it.
            var repeatingItemClone = firstRepeatingItem.CloneNode(true);
            var repeatingItemChildSdtElements =
                repeatingItemClone.Descendants<SdtElement>().Where(
                    e => !e.Ancestors<SdtElement>().Except(new[] { repeatingItemClone }).Any()).ToArray();
            content.AppendChild(repeatingItemClone);
            foreach (var sdt in repeatingItemChildSdtElements)
            {
                PopulateSdtElement(jObject, sdt, errorsCollector);
            }

            _logger.LogDebug("Placeholder row inserted successfully");
        }
    }

    private void ProcessText(SdtElement sdtElement, Tag tag, JToken token, ErrorsCollector errorsCollector)
    {
        _logger.LogDebug("{Tag}: {Token}", tag.Val, token);

        // Plain or Rich Text Content Control
        sdtElement.SdtProperties?.Elements<ShowingPlaceholder>().FirstOrDefault()?.Remove();
        // There are several possible types for <sdtContent> element (e.g. SdtContentBlock)
        // That's why we don't use a concrete type in the following line.
        var contentElement = FindContent(sdtElement);
        if (contentElement == null)
        {
            errorsCollector.AddError($"Placeholder {tag.Val} doesn't have any content area", sdtElement);
            _logger.LogWarning("Placeholder {Tag} doesn't have any content area", tag.Val);
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
                errorsCollector.AddError($"Placeholder '{tag.Val}' does not have a correct structure", sdtElement);
                _logger.LogWarning("Placeholder '{Tag}' does not have a correct structure", tag.Val);
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

    private void ProcessRichText(SdtElement sdtElement, Tag tag, JToken content, ErrorsCollector errorsCollector)
    {
        _logger.LogDebug("RichText: {Tag}: {Content}", tag.Val, content);

        if (sdtElement.Descendants<SdtContentText>().Any())
        {
            errorsCollector.AddError(
                $"HTML '{tag.Val}' cannot be written to PlainText control. Use Rich Text Control instead", sdtElement);
            _logger.LogWarning("HTML '{Tag}' cannot be written to PlainText control. Use Rich Text Control instead",
                tag.Val);
            return;
        }

        var contentElement = FindContent(sdtElement);
        if (contentElement == null)
        {
            errorsCollector.AddError($"Placeholder '{tag.Val}' doesn't have any content area", sdtElement);
            _logger.LogWarning("Placeholder '{Tag}' doesn't have any content area", tag.Val);
            return;
        }

        var mainDocumentPart = contentElement.Ancestors<Document>().First().MainDocumentPart;
        if (mainDocumentPart == null)
        {
            errorsCollector.AddError($"Failed to obtain MainDocumentPart while processing {tag.Val}", sdtElement);
            _logger.LogWarning("Failed to obtain MainDocumentPart while processing {Tag}", tag.Val);
            return;
        }

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