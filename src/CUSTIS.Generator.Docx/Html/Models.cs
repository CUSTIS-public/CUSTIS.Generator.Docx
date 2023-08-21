using AngleSharp.Dom;

namespace CUSTIS.Generator.Docx.Html;

public interface IToken
{
    public INode Node { get; }
}

public abstract class TagToken : IToken
{
    protected TagToken(INode node) => Node = node;

    public INode Node { get; }

    public string Name => Node.NodeName;

    public bool IsAnyOf(params string[] tags)
        => tags.Any(tag => Name.Equals(tag, StringComparison.InvariantCultureIgnoreCase));
}

public sealed class OpenTagToken : TagToken
{
    public OpenTagToken(INode node) : base(node)
    {
    }
}

public sealed class CloseTagToken : TagToken
{
    public CloseTagToken(INode node) : base(node)
    {
    }
}

public sealed class TextToken : IToken
{
    public TextToken(INode node) => Node = node;

    public INode Node { get; }

    public string Value => Node.TextContent;

    public bool IsWhiteSpace() => ((ReadOnlySpan<char>)Value).IsWhiteSpace();
}