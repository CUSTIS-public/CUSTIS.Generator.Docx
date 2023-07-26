using Newtonsoft.Json.Linq;

namespace CUSTIS.Generator.Docx;

public interface IDocumentProcessor
{
    void PopulateDocumentTemplate(Stream stream, JObject parameters, bool showErrorsInDocument = false);
}