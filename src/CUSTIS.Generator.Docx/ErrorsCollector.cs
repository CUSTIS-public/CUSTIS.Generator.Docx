using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CUSTIS.Generator.Docx;

internal record Error(string Message, OpenXmlElement Element);

internal class ErrorsCollector
{
    private const string Red = "C00000";

    public List<Error> Errors { get; } = new();

    public void AddError(string message, OpenXmlElement element)
    {
        Errors.Add(new(message, element));
    }

    public void AddErrors(MainDocumentPart mainPart)
    {
        if (!Errors.Any())
        {
            return;
        }

        var errorNum = 1;
        var bookmarkId = GetFreeBookmarkId(mainPart);

        var lastError = GenerateErrorsHeading();
        mainPart.Document.Body!.PrependChild(lastError);

        foreach (var error in Errors)
        {
            var bookmarkName = $"__docx_error_{errorNum++}";
            var newChild = GenerateError(error, bookmarkName);
            lastError.InsertAfterSelf(newChild);
            lastError = newChild;

            if (error.Element is not SdtElement sdtElement)
            {
                continue;
            }

            MakeRed(sdtElement);

            sdtElement.InsertBeforeSelf(new BookmarkStart
            {
                Id = new StringValue(bookmarkId.ToString()),
                Name = new StringValue(bookmarkName)
            });
            sdtElement.InsertAfterSelf(new BookmarkEnd
            {
                Id = new StringValue(bookmarkId.ToString())
            });

            bookmarkId++;
        }

        var space = new Paragraph(new Run(new Text()));
        lastError.InsertAfterSelf(space);
    }
    
    private static void MakeRed(SdtElement sdtElement)
    {
        foreach (var runProps in sdtElement.Descendants<RunProperties>())
        {
            runProps.PrependChild(new Color() { Val = Red });
            runProps.PrependChild(new Bold());
        }
    }

    private static Paragraph GenerateError(Error error, string bookmarkName)
    {
        var underline = new Underline
        {
            Val = new EnumValue<UnderlineValues>(UnderlineValues.Single)
        };
        var blueColor = new Color() { Val = "2F5496" };
        var hyperlink = new Hyperlink(new Run(new RunProperties(underline, blueColor), new Text(error.Message)))
        {
            Anchor = new StringValue(bookmarkName),
            History = new OnOffValue(true)
        };
        var newChild = new Paragraph(hyperlink);
        return newChild;
    }

    private static int GetFreeBookmarkId(MainDocumentPart mainPart)
    {
        return mainPart.Document.Descendants<BookmarkStart>().Select(s =>
        {
            if (int.TryParse(s.Id, out var i))
            {
                return (int?)i;
            }

            return null;
        }).Max() ?? 0 + 1;
    }

    private static OpenXmlElement GenerateErrorsHeading()
    {
        var bold = new Bold();
        var color = new Color() { Val = Red };
        var size = new FontSize() { Val = "36" };
        var text = new Text("Some errors occured while generating document");
        var runPro = new RunProperties(bold, color, size, text);
        OpenXmlElement lastError = new Paragraph(new Run(runPro));
        return lastError;
    }
}