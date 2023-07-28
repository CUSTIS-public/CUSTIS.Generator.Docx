using System;
using System.IO;
using ApprovalTests.Core;
using ApprovalTests.Reporters;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CUSTIS.Generator.Docx.Tests;

/// <summary>
/// Реализация <see cref="IApprovalFailureReporter"/> для работы в условиях отсутствия DiffTool в окружении
/// </summary>
public class CustomDiffReporter : IApprovalFailureReporter
{
    /// <summary>
    /// Показывает различия между <paramref name="approved"/> и <paramref name="received"/>.
    /// Если это возможно - через DiffTool, если нет - через <see cref="Assert.AreEqual(object,object)"/>
    /// </summary>
    /// <param name="approved"></param>
    /// <param name="received"></param>
    public void Report(string approved, string received)
    {
        if (Environment.GetEnvironmentVariable("GITLAB_CI") != null)
        {
            AssertFilesAreEqual(approved, received);
            return;
        }

        try
        {
            DiffReporter.INSTANCE.Report(approved, received);
        }
        catch (Exception e)
        {
            if (e.Message.Contains("Could not find a diff tool for extension:"))
            {
                AssertFilesAreEqual(approved, received);
            }
            else
            {
                throw;
            }
        }
    }

    private static void AssertFilesAreEqual(string approved, string received)
    {
        var approvedText = File.ReadAllText(approved);
        var receivedText = File.ReadAllText(received);
        Assert.AreEqual(approvedText, receivedText);
    }
}
