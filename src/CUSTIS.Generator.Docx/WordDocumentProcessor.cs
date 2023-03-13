using System.Diagnostics.CodeAnalysis;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json.Linq;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;

namespace CUSTIS.Generator.Docx;

public class WordDocumentProcessor : IDocumentProcessor
{
    private const string Wordml2006Ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private readonly ILogger<WordDocumentProcessor> _logger;

    private readonly ProcessorOptions _processorOptions;

    private static readonly XNamespace W = Wordml2006Ns;

    /// <summary> Controls (SDT) that we don't process  </summary>
    private static readonly Type[] UnsupportedSdtTypes = new[]
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

    public WordDocumentProcessor(ILogger<WordDocumentProcessor> log, ProcessorOptions? processorOptions = null)
    {
        _logger = log;
        _processorOptions = processorOptions ?? new ProcessorOptions();
    } 

    public async Task<MemoryStream> PopulateDocumentTemplate(string fileName, string jsonData,
        bool showErrorsInDocument = false)
    {
        await using var fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
        var filledStream = new MemoryStream();
        await fileStream.CopyToAsync(filledStream);

        var input = JObject.Parse(jsonData);

        PopulateDocumentTemplate(filledStream, input, showErrorsInDocument);
        filledStream.Position = 0;
        return filledStream;
    }

    public async Task<MemoryStream> PopulateDocumentTemplate(Stream template, object json,
        JsonSerializerOptions? jsonSerializerOptions = null,
        bool showErrorsInDocument = false)
    {
        var str = JsonSerializer.Serialize(json, jsonSerializerOptions ?? new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
        return await PopulateDocumentTemplate(template, str, showErrorsInDocument);
    }

    public async Task<MemoryStream> PopulateDocumentTemplate(Stream template, string json,
        bool showErrorsInDocument = false)
    {
        if (template == null) throw new ArgumentException(nameof(template));
        var input = JObject.Parse(json);
        var filledStream = new MemoryStream();
        await template.CopyToAsync(filledStream);

        PopulateDocumentTemplate(filledStream, input, showErrorsInDocument);
        filledStream.Position = 0;
        return filledStream;
    }

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
            = mainPart.Document.Descendants<SdtElement>().Where(e => !e.Ancestors<SdtElement>().Any()).ToArray();
        // Note: An <sdt> can be SdtBlock, SdtCell, SdtRow, SdtRun, SdtRunRub, and they are inherit
        // from SdtElement. That's why we used SdtElement above.
        foreach (var sdtElement in topLevelSdtElements)
        {
            PopulateSdtElement(parameters, sdtElement, errors);
        }

        if (showErrorsInDocument)
        {
            errors.AddErrors(mainPart);
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

        if (!TryGetTag(sdtElement, errorsCollector, out var tag))
        {
            return;
        }

        _logger.LogDebug("Processing tag: '{Tag}'", tag);
        
        var sdtProperties = sdtElement.SdtProperties;
        if (sdtProperties == null)
        {
            errorsCollector.AddError(
                $"Placeholder '{tag}' found without sdt properties. Placeholders without sdt properties are not supported",
                sdtElement);
            _logger.LogWarning(
                "Placeholder '{Tag}' found without sdt properties. Placeholders without sdt properties are not supported",
                tag);
            return;
        }

        var unsupportedDescendants =
            sdtProperties.Descendants().Where(d => UnsupportedSdtTypes.Contains(d.GetType())).ToArray();
        if (unsupportedDescendants.Any())
        {
            var typeNames = unsupportedDescendants.Select(d => d.GetType().Name);
            _logger.LogWarning(
                "Placeholder '{Tag}' has incorrect type '{Type}'. Only plain text, rich text and repeating section are supported",
                tag, typeNames);
            errorsCollector.AddError(
                $"Placeholder '{tag}' has incorrect type '{string.Join(", ", typeNames)}'. Only plain text, rich text and repeating section are supported",
                sdtElement);
            return;
        }

        if (tag.StartsWith("visible:", StringComparison.InvariantCultureIgnoreCase))
        {
            if (IsRepeatingSection(sdtProperties))
            {
                _logger.LogWarning(
                    "Conditional tag '{Tag}' can be applied only on plain or rich text, but is applied on repeating section",
                    tag);
                errorsCollector.AddError(
                    $"Conditional tag '{tag}' can be applied only on plain or rich text, but is applied on repeating section",
                    sdtElement);
                return;
            }

            // this tag controls visibility of contained element
            // "visible: x == y" -> "x == y"
            var visibilityCondition = tag.Replace("visible:", string.Empty, StringComparison.InvariantCultureIgnoreCase)
                .Trim();

            ProcessIfNecessary(sdtElement, visibilityCondition, parameters, errorsCollector);
            return;
        }

        if(!parameters.TryGetData(tag, out var token, out var dataError))
        {
            errorsCollector.AddError($"An error '{dataError}' occurred while fetching data from '{tag}'", sdtElement);
            _logger.LogInformation("An error '{DataError}' occurred while fetching data from '{Tag}'", dataError, tag);
            return;
        }

        if (token == null)
        {
            errorsCollector.AddError($"No data matched placeholder '{tag}'", sdtElement);
            _logger.LogInformation("No data matched placeholder '{Tag}'", tag);
            return;
        }

        try
        {
            if (sdtProperties.Descendants<SdtContentText>().Any())
            {
                // if <w:text> is present, then this is plain text control
                ProcessText(sdtElement, tag, token, errorsCollector);
            }
            else if (IsRepeatingSection(sdtProperties))
            {
                // if <w15:repeatingSection> is present, then this is repeating section
                ProcessRepeatingSection(sdtElement, tag, token, errorsCollector);
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
            errorsCollector.AddError($"An error '{e.Message}' occured while processing '{tag}'", sdtElement);
            _logger.LogError(e, "An error '{Message}' occured while processing '{Tag}'", e.Message, tag);
        }
    }

    private static bool IsRepeatingSection(SdtProperties sdtProperties)
    {
        return sdtProperties.Descendants<SdtRepeatedSection>().Any();
    }

    private void ProcessIfNecessary(SdtElement sdtElement,
        string visibilityCondition, JObject parameters, ErrorsCollector errorsCollector)
    {
        var success = visibilityCondition.TryEvaluate(parameters, out var visible, out var error);
        if (!success)
        {
            errorsCollector.AddError($"Failed to evaluate visibility condition '{visibilityCondition}'. Error: '{error}'",
                sdtElement);
            _logger.LogWarning("Failed to evaluate visibility condition '{VisibilityCondition}'. Error: '{Error}'",
                visibilityCondition, error);
            visible = true;
        }

        if (visible)
        {
            var sdtChildren = GetChildSdtElements(sdtElement);
            foreach (var sdt in sdtChildren)
            {
                PopulateSdtElement(parameters, sdt, errorsCollector);
            }
        }
        else
        {
            if (sdtElement.Parent is TableCell cell)
            {
                sdtElement.InsertAfterSelf(new Paragraph());
            }
            sdtElement.Remove();
        }
    }

    private bool TryGetTag(SdtElement sdtElement, ErrorsCollector errorsCollector,
        [NotNullWhen(true)] out string? tagVal)
    {
        var sdtProperties = sdtElement.SdtProperties;
        var tag = sdtProperties?.Descendants<Tag>().FirstOrDefault();
        tagVal = null;
        if (sdtProperties == null || tag == null)
        {
            errorsCollector.AddError("Placeholder found without tag. Placeholders without tag are not supported",
                sdtElement);
            _logger.LogWarning("Placeholder found without tag. Placeholders without tag are not supported");
            return false;
        }

        tagVal = tag.Val?.Value;
        if (string.IsNullOrWhiteSpace(tagVal))
        {
            errorsCollector.AddError("Placeholder found with an empty tag. Placeholders without tag are not supported",
                sdtElement);
            _logger.LogWarning("Placeholder found with an empty tag. Placeholders without tag are not supported");
            return false;
        }

        return true;
    }

    private void ProcessRepeatingSection(SdtElement sdtElement, string tag, JToken token,
        ErrorsCollector errorsCollector)
    {
        // Repeating Section Content Control
        if (token is not JArray tokens)
        {
            errorsCollector.AddError(
                $"The value of '{tag}' parameter is not an array. Parameter mapped to a repeating section can only be an array",
                sdtElement);
            _logger.LogWarning(
                "The value of '{Tag}' parameter is not an array. Parameter mapped to a repeating section can only be an array",
                tag);
            return;
        }

        var content = sdtElement.Elements().FirstOrDefault(e => e.XName == W + "sdtContent");
        if (content?.FirstChild == null)
        {
            errorsCollector.AddError(
                $"Encountered repeating '{tag}' with no content area for repeating item. It will be skipped",
                sdtElement);
            _logger.LogWarning(
                "Encountered repeating '{Tag}' with no content area for repeating item. It will be skipped",
                tag);
        }

        if (content!.FirstChild is not SdtElement firstRepeatingItem)
        {
            errorsCollector.AddError(
                $"Encountered repeating '{tag}' with wrong element instead of repeating item. It will be skipped",
                sdtElement);
            _logger.LogWarning(
                "Encountered repeating '{Tag}' with wrong element instead of repeating item. It will be skipped",
                tag);
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
                _logger.LogWarning(
                    "Element '{Token}' in '{Tag}' is not an object. It should be an object in '{{}}' braces",
                    tokenChild, tag);
                errorsCollector.AddError(
                    $"Element '{tokenChild}' in '{tag}' is not an object. It should be an object in '{{}}' braces",
                    sdtElement);
                continue;
            }

            // Find the first repeating section item, and clone it.
            var repeatingItemClone = firstRepeatingItem.CloneNode(true);
            var repeatingItemChildSdtElements = GetChildSdtElements(repeatingItemClone);
            content.AppendChild(repeatingItemClone);
            foreach (var sdt in repeatingItemChildSdtElements)
            {
                PopulateSdtElement(jObject, sdt, errorsCollector);
            }

            _logger.LogDebug("Placeholder row inserted successfully");
        }
    }

    private static SdtElement[] GetChildSdtElements(OpenXmlElement xmlElement)
    {
        var parents = new[] { xmlElement }.Concat(xmlElement.Ancestors<SdtElement>()).ToArray();
        // find only 'first level' descendants: we are not interested in descendants of descendants
        return xmlElement.Descendants<SdtElement>().Where(e => !e.Ancestors<SdtElement>().Except(parents).Any()).ToArray();
    }

    private void ProcessText(SdtElement sdtElement, string tag, JToken token, ErrorsCollector errorsCollector)
    {
        _logger.LogDebug("{Tag}: {Token}", tag, token);

        // Plain or Rich Text Content Control
        sdtElement.SdtProperties?.Elements<ShowingPlaceholder>().FirstOrDefault()?.Remove();
        // There are several possible types for <sdtContent> element (e.g. SdtContentBlock)
        // That's why we don't use a concrete type in the following line.
        var contentElement = FindContent(sdtElement);
        if (contentElement == null)
        {
            errorsCollector.AddError($"Placeholder {tag} doesn't have any content area", sdtElement);
            _logger.LogWarning("Placeholder {Tag} doesn't have any content area", tag);
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
        var strTokensTags = _processorOptions.ReplaceLineBreakWithTag
            ? SplitToken(strToken)
            : null;

        if (firstText == null)
        {
            if (firstRun != null)
            {
                if (strTokensTags == null)
                {
                    firstRun.AddChild(new Text(strToken));
                }
                else
                {
                    firstRun.AddChild(new Run(strTokensTags));
                }
            }
            else if (paragraph != null)
            {
                if (strTokensTags == null)
                {
                    paragraph.AddChild(new Run(new Text(strToken)));
                }
                else
                {
                    paragraph.AddChild(new Run(strTokensTags));
                }
            }
            else
            {
                errorsCollector.AddError($"Placeholder '{tag}' does not have a correct structure", sdtElement);
                _logger.LogWarning("Placeholder '{Tag}' does not have a correct structure", tag);
                return;
            }
        }
        else
        {
            if (strTokensTags == null)
            {
                firstText.Text = strToken;
            }
            else
            {
                firstText.Text = ((Text)strTokensTags[0]).Text;
                var prev = (OpenXmlElement)firstText;
                foreach (var t in strTokensTags.Skip(1))
                {
                    prev = firstText.Parent!.InsertAfter(t, prev);
                }
            }
            firstText.Parent!.Descendants<RunStyle>().FirstOrDefault(s => s.Val == "PlaceholderText")?.Remove();
        }

        paragraph?.Descendants<RunStyle>().FirstOrDefault(s => s.Val == "PlaceholderText")?.Remove();
        firstRun?.Descendants<RunStyle>().FirstOrDefault(s => s.Val == "PlaceholderText")?.Remove();
    }

    private static List<OpenXmlElement>? SplitToken(string token)
    {
        var strTokens = token
            .Replace(Environment.NewLine, "\n")
            .Replace("\r\n", "\n")
            .Replace("\n\r", "\n")
            .Replace("\r", "\n")
            .Split("\n");
        var strTokensTags = new List<OpenXmlElement>();
        if (strTokens.Length > 1)
        {
            foreach (var st in strTokens)
            {
                strTokensTags.Add(new Text(st));
                strTokensTags.Add(new Break());
            }
            strTokensTags.RemoveAt(strTokensTags.Count - 1);
        }
        return strTokensTags.Any() ? strTokensTags : null;
    }

    private static OpenXmlElement? FindContent(SdtElement sdtElement)
    {
        return sdtElement.Descendants().SingleOrDefault(e => e.XName == W + "sdtContent");
    }

    private void ProcessRichText(SdtElement sdtElement, string tag, JToken content, ErrorsCollector errorsCollector)
    {
        _logger.LogDebug("RichText: {Tag}: {Content}", tag, content);

        if (sdtElement.Descendants<SdtContentText>().Any())
        {
            errorsCollector.AddError(
                $"HTML '{tag}' cannot be written to PlainText control. Use Rich Text Control instead",
                sdtElement);
            _logger.LogWarning("HTML '{Tag}' cannot be written to PlainText control. Use Rich Text Control instead",
                tag);
            return;
        }

        var contentElement = FindContent(sdtElement);
        if (contentElement == null)
        {
            errorsCollector.AddError($"Placeholder '{tag}' doesn't have any content area", sdtElement);
            _logger.LogWarning("Placeholder '{Tag}' doesn't have any content area", tag);
            return;
        }

        var mainDocumentPart = contentElement.Ancestors<Document>().First().MainDocumentPart;
        if (mainDocumentPart == null)
        {
            errorsCollector.AddError($"Failed to obtain MainDocumentPart while processing {tag}", sdtElement);
            _logger.LogWarning("Failed to obtain MainDocumentPart while processing {Tag}", tag);
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