using docfill.Interfaces;
using docfill.Factory;
using docfill.Models;

class Program
{
    static void Main()
    {
        Console.WriteLine("Testing library...");

        var data = new InterviewFormModel
        {
            Name = "Alice",
            Email = "alice@example.com",
            Phone = "1234567890",
            Address = "123 Main Street",
            Table1 = new List<Table1Model>
            {
                new Table1Model { Id = 1, Name = "Alice Kumar", Year = 2020, Degree = "BSc" },
                new Table1Model { Id = 2, Name = "Rohit Sharma", Year = 2021, Degree = "BTech" },
                new Table1Model { Id = 3, Name = "Meera Thomas", Year = 2019, Degree = "BA" },
                new Table1Model { Id = 4, Name = "Sandeep Rao", Year = 2022, Degree = "MCA" },
            },
            Table2 = new List<Table2Model>
            {
                new Table2Model { Id = 5, Name = "Priya Nair", Year = 2018 },
                new Table2Model { Id = 6, Name = "Vikram Patel", Year = 2023 },
                new Table2Model { Id = 7, Name = "Anita Das", Year = 2020 },
                new Table2Model { Id = 8, Name = "Karan Mehta", Year = 2024 },
            }
        };

        var data2 = new JobApplicationModel
        {
            ApplicantName = "John Doe",
            Position = "Software Engineer",
            Email = "john.doe@example.com",
            Phone = "+1 555 123 4567",
            CoverLetter = "I am excited to apply for this position.",

            WorkExperience = new List<WorkExperienceModel>
            {
                new WorkExperienceModel
                {
                    Company = "TechCorp",
                    Position = "Backend Developer",
                    StartDate = "Jan 2021",
                    EndDate = "Dec 2022",
                    Description = "Worked on scalable API systems."
                },
                new WorkExperienceModel
                {
                    Company = "InnovateX",
                    Position = "Intern",
                    StartDate = "Jun 2020",
                    EndDate = "Dec 2020",
                    Description = "Assisted in microservice development."
                }
            },

            Education = new List<EducationModel>
            {
                new EducationModel
                {
                    Institution = "State University",
                    Degree = "BTech",
                    Field = "Computer Science",
                    GraduationYear = "2020"
                }
            }
        };

      
        IDocTemplate interviewTemplate = TemplateFactory.GetTemplate("InterviewForm");
        IDocTemplate jobApplicationTemplate = TemplateFactory.GetTemplate("JobApplication");


        //string baseUrl = "C:\\Users\\Admin\\Downloads\\docfillTest\\docfillTest\\docfillTest\\examples";

        string baseDir = AppContext.BaseDirectory;
        string projectDir = Path.GetFullPath(Path.Combine(baseDir, "..", "..", ".."));
        string baseUrl = Path.Combine(projectDir, "examples");

        interviewTemplate.FillTemplate(Path.Join(baseUrl, "interview_form.docx"), Path.Join(baseUrl, "interview_form_output.docx"), data);
        jobApplicationTemplate.FillTemplate(Path.Join(baseUrl, "job_application_form.docx"), Path.Join(baseUrl, "job_application_form_output.docx"), data2);


        Console.WriteLine("Done!");
    }
}
