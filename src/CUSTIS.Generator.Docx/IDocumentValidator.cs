namespace CUSTIS.Generator.Docx;

public interface IDocumentValidator
{
    bool CanProcessDocument(Stream stream);
}
