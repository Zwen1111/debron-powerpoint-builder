namespace DeBron.PowerPoint.Builder.Models;

public record Song
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public string Name { get; set; }
    public string Subtitle { get; set; }
    public string Lyrics { get; set; }
}