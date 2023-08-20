namespace CUSTIS.Generator.Docx.Html;

public interface IToken
{
}

public abstract class TagToken : IToken
{
    protected TagToken(string name)
    {
        Name = name;
    }

    public string Name { get; }
}

public sealed class OpenTagToken : TagToken
{
    public OpenTagToken(string name) : base(name)
    {
    }
}

public sealed class CloseTagToken : TagToken
{
    public CloseTagToken(string name) : base(name)
    {
    }
}

public sealed class TextToken : IToken
{
    public TextToken(string value)
    {
        Value = value;
    }

    public string Value { get; }

    public bool IsWhiteSpace() => ((ReadOnlySpan<char>)Value).IsWhiteSpace();
}