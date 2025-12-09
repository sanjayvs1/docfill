namespace docfill.Models;

/// <summary>
/// Data model for the Interview Form template.
/// </summary>
public class InterviewFormModel
{
    public string Name { get; set; } = string.Empty;
    public string Email { get; set; } = string.Empty;
    public string Phone { get; set; } = string.Empty;
    public string Address { get; set; } = string.Empty;

    public List<Table1Model> Table1 { get; set; } = new List<Table1Model>();
    public List<Table2Model> Table2 { get; set; } = new List<Table2Model>();
}

public class Table1Model
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
    public int Year { get; set; }
    public string Degree { get; set; } = string.Empty;
}

public class Table2Model
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
    public int Year { get; set; }
}
