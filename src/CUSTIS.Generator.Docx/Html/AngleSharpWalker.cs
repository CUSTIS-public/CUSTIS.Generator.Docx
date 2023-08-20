using AngleSharp.Dom;

namespace CUSTIS.Generator.Docx.Html;

public static class AngleSharpWalker
{
    /// <summary>
    /// Non-recursive implementation of DFS
    /// </summary>
    public static IEnumerable<(INode Node, AngleTagType TagType)> EnumerateTags(INode node)
    {
        var stack = new Stack<(INode Node, AngleTagType TagType)>();
        PushNode(node);
        do
        {
            var tag = stack.Pop();
            yield return tag;
            if (tag.TagType == AngleTagType.Open)
            {
                foreach (var childNode in tag.Node.ChildNodes.Reverse())
                {
                    PushNode(childNode);
                }
            }
        } while (stack.Count > 0);

        void PushNode(INode childNode)
        {
            stack.Push((childNode, AngleTagType.Close));
            stack.Push((childNode, AngleTagType.Open));
        }
    }
}