using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json.Linq;

namespace CUSTIS.Generator.Docx.Tests;

[TestClass]
public class ExpressionEvaluatorTests
{
    [DataTestMethod]
    [DataRow("null", false)]
    [DataRow("!null", true)]
    [DataRow("null == null", true)]
    [DataRow("null != null", false)]
    [DataRow("numField == 10", true)]
    [DataRow("numField != 10", false)]
    [DataRow("numField >= 10", true)]
    [DataRow("numField >= 11", false)]
    [DataRow("1 < numField", true)]
    [DataRow("100 < numField", false)]
    [DataRow("strField == 'str'", true)]
    [DataRow("strField == str", false)]
    [DataRow("strField == \"str\"", true)]
    [DataRow("strField == 'strf'", false)]
    [DataRow("noSuchFieldField", false)]
    [DataRow("noSuchFieldField == null", true)]
    [DataRow("emptyField == null", false)]
    [DataRow("emptyField == ''", true)]
    [DataRow("1 == 1", true)]
    [DataRow("1 != 1", false)]
    [DataRow("true", true)]
    [DataRow("false", false)]
    [DataRow("emptyField", false)]
    [DataRow("zeroField", false)]
    [DataRow("numField", true)]
    [DataRow("strField", true)]
    [DataRow("trueField", true)]
    [DataRow("!trueField", false)]
    [DataRow("falseField", false)]
    [DataRow("!falseField", true)]
    [DataRow("trueField != false", true)]
    [DataRow("trueField != true", false)]
    [DataRow("falseField != false", false)]
    [DataRow("falseField != true", true)]
    [DataRow("falseField != trueField", true)]
    [DataRow("falseField == trueField", false)]
    public void TestSuccess(string condition, bool expected)
    {
        //Arrange
        var input = JObject.Parse("{'numField': 10, zeroField: 0, 'trueField': true, 'falseField': false, 'strField': 'str', 'emptyField': ''}");
        
        //Act
        var success = condition.TryEvaluate(input, out var result, out var error);

        //Assert
        Assert.IsTrue(success);
        Assert.AreEqual(expected, result);
        Assert.IsNull(error);
    }
    
    [DataTestMethod]
    [DataRow("'1' < numField", "Operator '<' is allowed only for ints, but operands have types 'String' and 'Int32'")]
    [DataRow("strField < 'str'", "Operator '<' is allowed only for ints, but operands have types 'String' and 'String'")]
    [DataRow("!", "Operand is null or empty")]
    [DataRow("", "Operand is null or empty")]
    [DataRow("1 ==", "Right operand is null or empty")]
    [DataRow("== 1", "Left operand is null or empty")]
    public void TestErrors(string condition, string expectedError)
    {
        //Arrange
        var input = JObject.Parse("{'numField': 10, 'trueField': true, 'strField': 'str', emptyField: ''}");
        
        //Act
        var success = condition.TryEvaluate(input, out var result, out var error);

        //Assert
        Assert.IsFalse(success);
        Assert.AreEqual(false, result);
        Assert.AreEqual(expectedError, error);
    }
}