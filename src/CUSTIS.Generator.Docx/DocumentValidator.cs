using DocumentFormat.OpenXml.Packaging;

namespace CUSTIS.Generator.Docx;

public sealed class DocumentValidator : IDocumentValidator
{
    public bool CanProcessDocument(Stream stream)
    {
        try
        {
            using var doc = WordprocessingDocument.Open(stream, isEditable: false);
            var mainPart = doc.MainDocumentPart;
            return mainPart != null;
        }
        catch(Exception ex) when (ex is FileFormatException or OpenXmlPackageException)
        {
            return false;
        }
    }
}