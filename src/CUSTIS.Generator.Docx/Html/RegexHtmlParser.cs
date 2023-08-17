using System.Text.RegularExpressions;
using System.Web;

namespace CUSTIS.Generator.Docx.Html;

internal class RegexHtmlParser : IHtmlParser
{
    public IEnumerable<IToken> GetTokens(string html)
    {
        // разбивает строку на подстроки. Каждая подстрока либо тег (<...>), либо текст от тега до тега (>...<)
        // <p>my <div>string</div></p> будет развита на 6 подстрок:
        // <p>; my ; <div>; string; </div>; </p>
        var tokenizer = new Regex("<.+?>|[^<>]+");

        // получает название тега (будет в первом результате)
        // "< p font=>" ----> "p", "font"
        // "< / w:rPr>" ----> "/ w:rPr"
        var tagName = new Regex("\\/.*?[a-zA-Z:]+|[a-zA-Z:]+");

        static bool IsWhiteSpaceToken(Match token) => token.ValueSpan.IsWhiteSpace();

        static bool IsTagToken(Match token) => token.ValueSpan.StartsWith("<");

        static bool IsTextToken(Match token) => !IsWhiteSpaceToken(token) && !IsTagToken(token);

        bool IsBadTagToken(Match token) => !tagName.IsMatch(token.Value);

        string? IsCloseTagToken(Match token)
        {
            var match = tagName.Match(token.Value);
            return match.ValueSpan.StartsWith("/")
                ? match.ValueSpan.Slice(1).Trim().ToString()
                : null;
        }

        bool IsBadCloseTagToken(Match token) => tagName.Match(token.Value).ValueSpan.Length == 1;

        foreach (Match match in tokenizer.Matches(HttpUtility.HtmlDecode(html)))
        {
            if (IsTagToken(match) && (IsBadTagToken(match) || IsCloseTagToken(match) is { } && IsBadCloseTagToken(match)))
            {
                continue;
            }

            if (IsWhiteSpaceToken(match))
                yield return new WhiteSpaceToken(match.Value);
            else if (IsTagToken(match))
                if (IsCloseTagToken(match) is { } name)
                    yield return new CloseTagToken(name);
                else
                    yield return new OpenTagToken(tagName.Match(match.Value).Value);
            else if (IsTextToken(match))
                yield return new TextToken(match.Value);
        }
    }
}