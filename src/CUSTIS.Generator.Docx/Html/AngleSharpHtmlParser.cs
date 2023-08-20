using AngleSharp.Dom;
using AngleSharp.Html.Parser;

namespace CUSTIS.Generator.Docx.Html;

internal class AngleSharpHtmlParser : IHtmlParser
{
    public IEnumerable<IToken> GetTokens(string html)
        => new HtmlParser().ParseDocument(html).Body is { } body
            ? AngleSharpWalker.EnumerateTags(body)
                .Select(tag => ToToken(tag.Node, tag.TagType))
                .OfType<IToken>()
            : Enumerable.Empty<IToken>();

    private static IToken? ToToken(INode node, AngleTagType tagType)
        => node.NodeType switch
        {
            NodeType.Text when tagType is AngleTagType.Open => ((ReadOnlySpan<char>)node.TextContent).IsWhiteSpace()
                ? new WhiteSpaceToken(node.TextContent)
                : new TextToken(node.TextContent),
            NodeType.Element when tagType is AngleTagType.Open => new OpenTagToken(node.NodeName),
            NodeType.Element when tagType is AngleTagType.Close => new CloseTagToken(node.NodeName),
            _ => null
        };
}