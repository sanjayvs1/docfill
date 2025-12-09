namespace docfill.Models;

/// <summary>
/// Data model for the Job Application template.
/// </summary>
public class JobApplicationModel
{
    public string ApplicantName { get; set; } = string.Empty;
    public string Position { get; set; } = string.Empty;
    public string Email { get; set; } = string.Empty;
    public string Phone { get; set; } = string.Empty;
    public string CoverLetter { get; set; } = string.Empty;
    public List<WorkExperienceModel> WorkExperience { get; set; } = new List<WorkExperienceModel>();
    public List<EducationModel> Education { get; set; } = new List<EducationModel>();
}

public class WorkExperienceModel
{
    public string Company { get; set; } = string.Empty;
    public string Position { get; set; } = string.Empty;
    public string StartDate { get; set; } = string.Empty;
    public string EndDate { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
}

public class EducationModel
{
    public string Institution { get; set; } = string.Empty;
    public string Degree { get; set; } = string.Empty;
    public string Field { get; set; } = string.Empty;
    public string GraduationYear { get; set; } = string.Empty;
}

