using docfill.Interfaces;
using docfill.Models;
using docfill.Utils;

namespace docfill.Templates;

/// <summary>
/// Template implementation for Job Application documents.
/// </summary>
public class JobApplicationTemplate : IDocTemplate
{
    public string TemplateType => "JobApplication";

    public void FillTemplate(string templateFilePath, string outputFilePath, object data)
    {
        if (data is not JobApplicationModel model)
            throw new ArgumentException($"Expected {nameof(JobApplicationModel)}, got {data.GetType().Name}");

        // Copy template to output location
        FileHelper.CopyTemplate(templateFilePath, outputFilePath);

        // Prepare text replacements
        var replacements = new Dictionary<string, string>
        {
            { "{{APPLICANT_NAME}}", model.ApplicantName },
            { "{{POSITION}}", model.Position },
            { "{{EMAIL}}", model.Email },
            { "{{PHONE}}", model.Phone },
            { "{{COVER_LETTER}}", model.CoverLetter }
        };

        // Fill text placeholders
        FileHelper.ProcessDocument(outputFilePath, main =>
        {
            DocxHelper.ReplaceText(main, replacements);
        });

        // Convert work experience to table format
        var workExperienceData = model.WorkExperience.Select(we => new List<string>
        {
            we.Company,
            we.Position,
            we.StartDate,
            we.EndDate,
            we.Description
        }).ToList();

        // Convert education to table format
        var educationData = model.Education.Select(ed => new List<string>
        {
            ed.Institution,
            ed.Degree,
            ed.Field,
            ed.GraduationYear
        }).ToList();

        // Fill tables (if they exist in the template)
        FileHelper.ProcessDocument(outputFilePath, main =>
        { 
                DocxHelper.FillTable(main, "work_experience", workExperienceData);
                DocxHelper.FillTable(main, "education", educationData);
        });
    }
}

