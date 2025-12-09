using docfill.Interfaces;
using docfill.Models;
using docfill.Utils;

namespace docfill.Templates;

/// <summary>
/// Template implementation for Interview Form documents.
/// </summary>
public class InterviewFormTemplate : IDocTemplate
{
    public string TemplateType => "InterviewForm";

    public void FillTemplate(string templateFilePath, string outputFilePath, object data)
    {
        if (data is not InterviewFormModel model)
            throw new ArgumentException($"Expected {nameof(InterviewFormModel)}, got {data.GetType().Name}");

        // Copy template to output location
        FileHelper.CopyTemplate(templateFilePath, outputFilePath);

        // Prepare text replacements
        var replacements = new Dictionary<string, string>
        {
            { "{{NAME}}", model.Name },
            { "{{EMAIL}}", model.Email },
            { "{{PHONE}}", model.Phone },
            { "{{ADDRESS}}", model.Address }
        };

        // Fill text placeholders
        FileHelper.ProcessDocument(outputFilePath, main =>
        {
            DocxHelper.ReplaceText(main, replacements);
        });

        // Convert table models to List<List<string>> format
        var table1Data = model.Table1.Select(t => new List<string>
        {
            t.Id.ToString(),
            t.Name,
            t.Year.ToString(),
            t.Degree
        }).ToList();

        var table2Data = model.Table2.Select(t => new List<string>
        {
            t.Id.ToString(),
            t.Name,
            t.Year.ToString()
        }).ToList();

        // Fill tables
        FileHelper.ProcessDocument(outputFilePath, main =>
        {
            DocxHelper.FillTable(main, "table_1", table1Data);
            DocxHelper.FillTable(main, "table_2", table2Data);
        });

        long width = 200 * 9525;
        long height = 200 * 9525;

        string baseDir = AppContext.BaseDirectory;
        string projectDir = Path.GetFullPath(Path.Combine(baseDir, "..", "..", ".."));

        // Replace Image
        FileHelper.ProcessDocument(outputFilePath, main =>
        {
            DocxHelper.AddImage(main, "{{IMAGE_1}}", Path.Join(projectDir, "sample.png"), width, height);
        });


    }
}

