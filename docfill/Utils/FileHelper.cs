using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace docfill.Utils;

/// <summary>
/// Helper methods for file operations (loading, saving DOCX files).
/// </summary>

public static class FileHelper
{
    /// <summary>
    /// Copies a template file to the output location.
    /// </summary>
    public static void CopyTemplate(string templateFilePath, string outputFilePath)
    {
        if (!File.Exists(templateFilePath))
            throw new FileNotFoundException($"Template file not found: {templateFilePath}");

        File.Copy(templateFilePath, outputFilePath, overwrite: true);
    }

    /// <summary>
    /// Opens a DOCX document and returns the main document part's body.
    /// </summary>
    public static Body OpenDocumentBody(string filePath)
    {
        using var document = WordprocessingDocument.Open(filePath, true);
        var main = document.MainDocumentPart;
        if (main == null)
            throw new InvalidOperationException("MainDocumentPart not found in document.");

        var doc = main.Document;
        if (doc == null)
            throw new InvalidOperationException("Document XML part not found.");

        var body = doc.Body;
        if (body == null)
            throw new InvalidOperationException("Body element not found.");

        return body;
    }

    /// <summary>
    /// Executes an action on a DOCX document body and saves the document.
    /// </summary>
    public static void ProcessDocument(string filePath, Action<MainDocumentPart> action)
    {
        using var document = WordprocessingDocument.Open(filePath, true);
        var main = document.MainDocumentPart ?? throw new InvalidOperationException("MainDocumentPart missing");
        action(main);
        main.Document.Save();
    }

}

