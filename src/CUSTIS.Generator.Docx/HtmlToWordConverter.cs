﻿using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;

namespace CUSTIS.Generator.Docx;

public static class HtmlToWordConverter
{
    public record ConvertResult(IList<Paragraph> Paragraphs, IList<AbstractNum> AbstractNums,
        IList<NumberingInstance> NumberingInstances);

    private record ListInfo(NumberingInstance Format)
    {
        public int Level { get; set; } = 0;
    }

    public static ConvertResult ConvertToDocx(this string htmlText, MainDocumentPart existingDoc)
    {
        var result = new ConvertResult(new List<Paragraph>(), new List<AbstractNum>(), new List<NumberingInstance>());

        var paragraphs = result.Paragraphs;

        // разбивает строку на подстроки. Каждая подстрока либо тег (<...>), либо текст от тега до тега (>...<)
        // <p>my <div>string</div></p> будет развита на 6 подстрок:
        // <p>; my ; <div>; string; </div>; </p>
        var tokenizer = new Regex("<.+?>|[^<>]+");

        // получает название тега (будет в первом результате)
        // "< p font=>" ----> "p", "font"
        // "< / w:rPr>" ----> "/ w:rPr"
        var tagName = new Regex("\\/.*?[a-zA-Z:]+|[a-zA-Z:]+");
        var current = new StringBuilder();

        ListInfo? currentList = null;
        foreach (var token in GetTokens(htmlText))
        {
            if (IsWhiteSpaceToken(token))
            {
                //пробельный текст
                if (current.Length <= 0 || current[^1] != ' ')
                {
                    current.Append(' ');
                }
            }
            else if (IsTagToken(token))
            {
                if (IsCloseTagToken(token) is { } closingTag)
                {
                    //закрывающийся тег
                    if (currentList != null && IsAnyTagOf(closingTag, "ul", "ol"))
                    {
                        AppendParagraph(paragraphs, current, currentList);
                        current = new StringBuilder();

                        currentList.Level--;
                        if (currentList.Level < 0)
                        {
                            currentList = null;
                        }
                    }
                }
                else
                {
                    //открывающийся тег
                    if (IsAnyTagOfToken(token, "p", "li", "br", "br/"))
                    {
                        AppendParagraph(paragraphs, current, currentList);
                        current = new StringBuilder();
                    }

                    var isBulletList = IsAnyTagOfToken(token, "ul");
                    var isNumberedList = IsAnyTagOfToken(token, "ol");
                    if (isBulletList || isNumberedList)
                    {
                        AppendParagraph(paragraphs, current, currentList);
                        current = new StringBuilder();

                        if (currentList == null)
                        {
                            var listFormat = CreateList(existingDoc, result, isBulletList ? NumberFormatValues.Bullet : NumberFormatValues.Decimal);
                            currentList = new(listFormat);
                        }
                        else
                        {
                            currentList.Level++;
                        }
                    }
                }
            }
            else
            {
                //текст
                current.Append(token.ValueSpan);
            }
        }

        AppendParagraph(paragraphs, current, currentList);

        return result;

        static bool IsWhiteSpaceToken(Token token) => token.ValueSpan.IsWhiteSpace();

        static bool IsTagToken(Token token) => token.ValueSpan.StartsWith("<");

        bool IsBadTagToken(Token token) => !tagName.IsMatch(token.Value);

        string? IsCloseTagToken(Token token)
        {
            var match = tagName.Match(token.Value);
            return match.ValueSpan.StartsWith("/")
                ? match.ValueSpan.Slice(1).Trim().ToString()
                : null;
        }

        bool IsBadCloseTagToken(Token token) => tagName.Match(token.Value).ValueSpan.Length == 1;

        bool IsAnyTagOfToken(Token token, params string[] tags)
            => IsAnyTagOf(tagName.Match(token.Value).Value, tags);

        bool IsAnyTagOf(string tagName, params string[] tags)
            => tags.Any(tag => tagName.Equals(tag, StringComparison.InvariantCultureIgnoreCase));

        IEnumerable<Token> GetTokens(string html)
        {
            foreach (Match match in tokenizer.Matches(HttpUtility.HtmlDecode(html)))
            {
                var token = new Token(match);
                if (IsTagToken(token) && (IsBadTagToken(token) || IsCloseTagToken(token) is { } && IsBadCloseTagToken(token)))
                {
                    continue;
                }

                yield return token;
            }
        }
    }

    public class Token
    {
        private readonly Match _match;

        public Token(Match match)
        {
            _match = match;
        }

        public ReadOnlySpan<char> ValueSpan => _match.ValueSpan;
        public string Value => _match.Value;
    }

    private static void AppendParagraph(IList<Paragraph> paragraphs, StringBuilder current, ListInfo? currentList)
    {
        var text = current.ToString().Trim();

        if (text.Length <= 0)
        {
            return;
        }

        if (currentList != null)
        {
            paragraphs.Add(CreateListItem(text, currentList));
        }
        else
        {
            paragraphs.Add(new Paragraph(new Run(new Text(text))));
        }
    }

    private static NumberingInstance CreateList(MainDocumentPart existingDoc, ConvertResult result,
        NumberFormatValues numberFormat)
    {
        var maxAbstractNumberId = result.AbstractNums.Max(x => x.AbstractNumberId?.Value)
                                  ?? NumberingDescendants<AbstractNum>(existingDoc)?.Max(n => n.AbstractNumberId?.Value)
                                  ?? -1;

        var maxNumberId = result.NumberingInstances.Max(x => x.NumberID?.Value)
                          ?? NumberingDescendants<NumberingInstance>(existingDoc)?.Max(n => n.NumberID?.Value)
                          ?? 0;

        Func<int, string> levelText;
        if (numberFormat == NumberFormatValues.Bullet)
        {
            var bullets = new[] { "•", "◦", "·" };
            levelText = l => bullets[l % bullets.Length];
        }
        else
        {
            levelText = l => $"%{l + 1}.";
        }

        var listAbstractFormat = CreateAbstractFormat(maxAbstractNumberId + 1, numberFormat, levelText);
        result.AbstractNums.Add(listAbstractFormat);
        var listFormat = CreateFormatInstance(listAbstractFormat, maxNumberId + 1);
        result.NumberingInstances.Add(listFormat);

        return listFormat;
    }

    private static IEnumerable<T>? NumberingDescendants<T>(MainDocumentPart existingDoc) where T : OpenXmlElement
    {
        return existingDoc.NumberingDefinitionsPart?.Numbering?.Descendants<T>();
    }

    private static NumberingInstance CreateFormatInstance(AbstractNum abstractFormat, int maxNumberId)
    {
        return new NumberingInstance(
            new AbstractNumId() { Val = abstractFormat.AbstractNumberId }) { NumberID = maxNumberId };
    }

    private static AbstractNum CreateAbstractFormat(int maxAbstractNumberId, NumberFormatValues numberFormat,
        Func<int, string> levelText)
    {
        var levels = Enumerable.Range(0, 9).Select(l =>
            new Level(
                new NumberingFormat() { Val = numberFormat },
                new LevelText() { Val = levelText(l) },
                new ParagraphProperties
                {
                    Indentation = new()
                    {
                        Left = $"{720 * (l + 1)}",
                        Hanging = "360"
                    }
                })
            {
                LevelIndex = l,
                StartNumberingValue = new() { Val = 1 }
            });
        return new AbstractNum(levels) { AbstractNumberId = maxAbstractNumberId };
    }

    private static Paragraph CreateListItem(string text, ListInfo listInfo)
    {
        var listItem = new Paragraph(new Run(new Text(text)));
        listItem.ParagraphProperties = new()
        {
            NumberingProperties = new()
            {
                NumberingLevelReference = new() { Val = listInfo.Level },
                NumberingId = new() { Val = listInfo.Format.NumberID }
            }
        };
        return listItem;
    }
}