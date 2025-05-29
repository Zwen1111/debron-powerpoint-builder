using System.Text.Json;
using System.Text.Json.Serialization;

namespace DeBron.PowerPoint.Builder.Models;

public interface IPresentationPart
{
    public Guid Id { get; set; }
}

public record Song : IPresentationPart
{
    public bool UseSubtitle { get; set; } = false;
    public Guid Id { get; set; } = Guid.NewGuid();
    public string Name { get; set; }
    public string Subtitle { get; set; }
    public string Lyrics { get; set; }
}

public record Collection : IPresentationPart
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public string FirstGoal { get; set; }
    public string SecondGoal { get; set; }
}

public record Prayer : IPresentationPart
{
    public Guid Id { get; set; } = Guid.NewGuid();
}

public record BibleReading : IPresentationPart
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public string Title { get; set; }
    public string Text { get; set; }
}

public record TrustAndGreeting : IPresentationPart
{
    public Guid Id { get; set; } = Guid.NewGuid();
}

public record ChildrenMoment : IPresentationPart
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public string Person { get; set; }
}

public class PresentationPartConverter : JsonConverter<IPresentationPart>
{
    public override IPresentationPart Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        using var jsonDoc = JsonDocument.ParseValue(ref reader);
        var root = jsonDoc.RootElement;

        var type = root.GetProperty("Type").GetString();
        return type switch
        {
            nameof(Song) => JsonSerializer.Deserialize<Song>(root.GetRawText(), options)!,
            nameof(Collection) => JsonSerializer.Deserialize<Collection>(root.GetRawText(), options)!,
            nameof(Prayer) => JsonSerializer.Deserialize<Prayer>(root.GetRawText(), options)!,
            nameof(BibleReading) => JsonSerializer.Deserialize<BibleReading>(root.GetRawText(), options)!,
            nameof(TrustAndGreeting) => JsonSerializer.Deserialize<TrustAndGreeting>(root.GetRawText(), options)!,
            nameof(ChildrenMoment) => JsonSerializer.Deserialize<ChildrenMoment>(root.GetRawText(), options)!,
            _ => throw new NotSupportedException($"Type '{type}' wordt niet ondersteund.")
        };
    }

    public override void Write(Utf8JsonWriter writer, IPresentationPart value, JsonSerializerOptions options)
    {
        var type = value.GetType().Name;
        var json = JsonSerializer.SerializeToElement(value, value.GetType(), options);

        using var doc = JsonDocument.Parse(json.GetRawText());
        writer.WriteStartObject();
        writer.WriteString("Type", type);

        foreach (var prop in doc.RootElement.EnumerateObject())
        {
            prop.WriteTo(writer);
        }

        writer.WriteEndObject();
    }
}


public enum SlideLayout
{
    WelkomVooraf,
    LiturgieVooraf,
    CollecteVooraf,
    Thema,
    Welkom,
    Liturgie,
    PaarsMetTitel,
    TussenSlide,
    LiedAankondiging,
    LiedAankondigingOverlay,
    Ondertiteling,
    Gebed,
    BlauwMetTitel,
    LuisterLiedAankondiging,
    WitMetLiedtekst,
    Koffermoment,
    BijbellezenAankondiging,
    Bijbeltekst,
    CollecteTweeDoelen,
    CollecteEenDoel,
    TotZiensMetGebed,
    TotZiens
}