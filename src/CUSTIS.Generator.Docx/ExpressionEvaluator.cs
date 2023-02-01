using System.Diagnostics.CodeAnalysis;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;

namespace CUSTIS.Generator.Docx;

internal static class ExpressionEvaluator
{
    /// <summary> Evaluates <paramref name="condition" />, using data from <paramref name="obj" /> </summary>
    /// <remarks>
    ///     Evaluates simple conditions like "A == B".
    ///     Supports: ==, !=, &lt;, &gt;, &lt;=, &gt;=
    /// </remarks>
    public static bool TryEvaluate(this string condition, JObject obj, out bool result,
        [NotNullWhen(false)] out string? error)
    {
        result = false;
        error = null;

        // x == y, x <= y, ...
        var expressionWithOperands = new Regex("(.*)(==|!=|<=|>=|<|>)(.*)");
        if (expressionWithOperands.IsMatch(condition))
        {
            var match = expressionWithOperands.Match(condition);
            var left = match.Groups[1].Value.Trim();
            var op = match.Groups[2].Value;
            var right = match.Groups[3].Value.Trim();

            return TryEvaluate(obj, left, op, right, ref result, ref error);
        }

        // !x
        var notExpression = new Regex("!(.*)");
        if (notExpression.IsMatch(condition))
        {
            var match = notExpression.Match(condition);
            var expression = match.Groups[1].Value.Trim();
            
            var evaluationResult = TryEvaluateBool(expression, obj, ref result, ref error);
            result = evaluationResult && !result;
            return evaluationResult;
        }

        return TryEvaluateBool(condition, obj, ref result, ref error);
    }

    private static bool TryEvaluateBool(string condition, JObject obj, ref bool result,
    [NotNullWhen(false)] ref string? error)
    {
        var expression = condition.Trim();
        if (string.IsNullOrEmpty(expression))
        {
            error = "Operand is null or empty";
            return false;
        }
        
        var evaluated = Evaluate(expression, obj);

        if (evaluated is bool boolCondition)
        {
            result = boolCondition;
            return true;
        }

        if (evaluated == null
            || evaluated is string strCondition && string.IsNullOrEmpty(strCondition)
            || evaluated is 0)
        {
            result = false;
            return true;
        }

        result = true;
        return true;
    }

    private static bool TryEvaluate(JObject obj, string left, string op, string right, ref bool result,
        [NotNullWhen(false)] ref string? error)
    {
        if (string.IsNullOrEmpty(left))
        {
            error = "Left operand is null or empty";
            return false;
        }
        
        if (string.IsNullOrEmpty(right))
        {
            error = "Right operand is null or empty";
            return false;
        }
        
        var leftObj = Evaluate(left, obj);
        var rightObj = Evaluate(right, obj);

        switch (op)
        {
            case "==":
                result = object.Equals(leftObj, rightObj);
                return true;
            case "!=":
                result = !object.Equals(leftObj, rightObj);
                return true;
        }

        if (leftObj is not int leftInt || rightObj is not int rightInt)
        {
            error = $"Operator '{op}' is allowed only for ints, but operands have types '{leftObj?.GetType().Name}' and '{rightObj?.GetType().Name}'";
            return false;
        }

        result = op switch
        {
            "<=" => leftInt <= rightInt,
            "<" => leftInt < rightInt,
            ">=" => leftInt >= rightInt,
            ">" => leftInt > rightInt,
            _ => false
        };
        return true;
    }

    private static object? Evaluate(string expression, JObject obj)
    {
        if (expression == string.Empty || expression.Equals("null", StringComparison.InvariantCultureIgnoreCase))
        {
            return null;
        }

        var token = obj.SelectToken(expression);
        if (token != null)
        {
            return token.Type switch
            {
                JTokenType.Integer => token.ToObject<int>(),
                JTokenType.Boolean => token.ToObject<bool>(),
                _ => token.ToString()
            };
        }

        if (int.TryParse(expression, out var intResult))
        {
            return intResult;
        }

        if (bool.TryParse(expression, out var boolResult))
        {
            return boolResult;
        }

        if (expression.StartsWith("'") && expression.EndsWith("'") && expression.Length >= 2)
        {
            return expression[1..^1];
        }

        if (expression.StartsWith('"') && expression.EndsWith('"') && expression.Length >= 2)
        {
            return expression[1..^1];
        }

        // trying to evaluate the field that is not presented in object
        return null;
    }
}