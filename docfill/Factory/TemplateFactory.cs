using docfill.Interfaces;
using docfill.Templates;

namespace docfill.Factory;

/// <summary>
/// Factory for creating template instances at runtime.
/// </summary>
public static class TemplateFactory
{
    private static readonly Dictionary<string, IDocTemplate> _templates = new()
    {
        { "InterviewForm", new InterviewFormTemplate() },
        { "JobApplication", new JobApplicationTemplate() }
    };

    /// <summary>
    /// Gets a template instance by type name.
    /// </summary>
    /// <param name="templateType">The template type identifier (e.g., "InterviewForm", "JobApplication").</param>
    /// <returns>An instance of the requested template.</returns>
    /// <exception cref="ArgumentException">Thrown when the template type is not found.</exception>
    public static IDocTemplate GetTemplate(string templateType)
    {
        if (!_templates.TryGetValue(templateType, out var template))
        {
            var availableTypes = string.Join(", ", _templates.Keys);
            throw new ArgumentException(
                $"Template type '{templateType}' not found. Available types: {availableTypes}");
        }

        return template;
    }

    /// <summary>
    /// Gets all available template types.
    /// </summary>
    public static IEnumerable<string> GetAvailableTemplateTypes()
    {
        return _templates.Keys;
    }

    /// <summary>
    /// Registers a new template type (useful for extensibility).
    /// </summary>
    public static void RegisterTemplate(IDocTemplate template)
    {
        _templates[template.TemplateType] = template;
    }
}

