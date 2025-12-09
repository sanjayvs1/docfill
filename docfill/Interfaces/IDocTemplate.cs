namespace docfill.Interfaces;

/// <summary>
/// Interface for document templates that can be filled with data.
/// </summary>
public interface IDocTemplate
{
    /// <summary>
    /// Gets the template type identifier.
    /// </summary>
    string TemplateType { get; }

    /// <summary>
    /// Fills the template with data and saves to the output path.
    /// </summary>
    /// <param name="templateFilePath">Path to the template DOCX file.</param>
    /// <param name="outputFilePath">Path where the filled document will be saved.</param>
    /// <param name="data">The data model to fill the template with.</param>
    void FillTemplate(string templateFilePath, string outputFilePath, object data);
}
