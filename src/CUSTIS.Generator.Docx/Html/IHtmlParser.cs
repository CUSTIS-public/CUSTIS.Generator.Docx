namespace CUSTIS.Generator.Docx.Html;

internal interface IHtmlParser
{
    IEnumerable<IToken> GetTokens(string html);
}