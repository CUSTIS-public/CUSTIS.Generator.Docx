using DocumentFormat.OpenXml;

namespace CUSTIS.Generator.Docx;

internal record Error(string Message, OpenXmlElement Element);

internal class ErrorsCollector
{
    public List<Error> Errors { get; } = new();

    public void AddError(string message, OpenXmlElement element)
    {
        Errors.Add(new(message, element));
    }
}