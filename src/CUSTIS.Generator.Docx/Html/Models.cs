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

public class OpenTagToken : TagToken
{
    public OpenTagToken(string name) : base(name)
    {
    }
}

public class CloseTagToken : TagToken
{
    public CloseTagToken(string name) : base(name)
    {
    }
}

public class TextToken : IToken
{
    public TextToken(string value)
    {
        Value = value;
    }

    public string Value { get; }
}

public sealed class WhiteSpaceToken : TextToken
{
    public WhiteSpaceToken(string value) : base(value)
    {
    }
}