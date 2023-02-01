using System.Diagnostics.CodeAnalysis;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace CUSTIS.Generator.Docx;

internal static class JObjectExtensions
{
    public static bool TryGetData(this JObject obj, string tag,
        out JToken? token, [NotNullWhen(false)] out string? error)
    {
        error = null;
        if (obj.TryGetValue(tag, out token))
        {
            return true;
        }

        try
        {
            token = obj.SelectToken(tag);
            return true;
        }
        catch (JsonException e) when (e.Message.Contains("Path returned multiple tokens"))
        {
            token = new JArray(obj.SelectTokens(tag));
            return true;
        }
        catch (Exception e)
        {
            token = null;
            error = e.Message;
            return false;
        }
    }
}