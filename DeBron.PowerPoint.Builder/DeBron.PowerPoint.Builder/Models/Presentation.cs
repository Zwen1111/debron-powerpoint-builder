namespace DeBron.PowerPoint.Builder.Models;

public class Presentation
{
    public string Theme { get; set; } = "Dit is het thema";
    public List<PresentationPart> Parts { get; init; } = new();
}