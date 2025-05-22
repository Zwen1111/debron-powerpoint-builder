namespace DeBron.PowerPoint.Builder.Models;

public class Song
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public string Name { get; set; }
    public string Artist { get; set; }
    public string Lyrics { get; set; }
}