using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CUSTIS.Generator.Docx.Tests;

[TestClass]
public class DocumentValidatorTests
{
    [DataTestMethod]
    public void CanProcessDocument_ValidDocument_ReturnsTrue()
    {
        //Arrange
        using FileStream fs = File.OpenRead(Path.Combine(@"Samples", "ComplexDocument.template.docx"));

        //Act
        var result = new DocumentValidator().CanProcessDocument(fs);

        //Assert
        Assert.IsTrue(result);
    }

    [DataTestMethod]
    public void CanProcessDocument_InvalidDocument_ReturnsFalse()
    {
        //Arrange
        using FileStream fs = File.OpenRead(Path.Combine(@"Samples", "ComplexDocument.input.json"));

        //Act
        var result = new DocumentValidator().CanProcessDocument(fs);

        //Assert
        Assert.IsFalse(result);
    }
}