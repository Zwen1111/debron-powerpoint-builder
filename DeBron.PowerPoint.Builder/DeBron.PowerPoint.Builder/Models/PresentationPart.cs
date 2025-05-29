using System.Text.Json;
using System.Text.Json.Serialization;

namespace DeBron.PowerPoint.Builder.Models;

public abstract record PresentationPart
{
    public Guid Id { get; set; } = Guid.NewGuid();
    internal readonly Dictionary<string, string> PlaceholderValues = new();

    public abstract IEnumerable<(SlideLayout Layout, Dictionary<string, string> PlaceholderValues)> GetSlides(bool isPriorToService = false);
}

public record Song : PresentationPart
{
    public bool UseSubtitle { get; set; } = false;

    public string Titel
    {
        get => PlaceholderValues.TryGetValue(nameof(Titel), out var value) ? value : string.Empty;
        set => PlaceholderValues[nameof(Titel)] = value;
    }

    public string Ondertitel
    {
        get => PlaceholderValues.TryGetValue(nameof(Ondertitel), out var value) ? value : string.Empty;
        set => PlaceholderValues[nameof(Ondertitel)] = value;
    }
    public string Liedtekst
    {
        get => PlaceholderValues.TryGetValue(nameof(Liedtekst), out var value) ? value : string.Empty;
        set => PlaceholderValues[nameof(Liedtekst)] = value;
    }

    public override IEnumerable<(SlideLayout Layout, Dictionary<string, string> PlaceholderValues)> GetSlides(bool isPriorToService = false)
    {
        yield return (SlideLayout.LiedAankondigingOverlay, PlaceholderValues);

        yield return (SlideLayout.Ondertiteling, new Dictionary<string, string>
        {
            { nameof(Liedtekst), string.Empty }
        });

        var lyricsPerSlide = Liedtekst.Trim().Split("\n\n")
            .SelectMany<string, string>(x => x.StartsWith("\n") ? [string.Empty, x.Trim()] : [x.Trim()]).ToList();

        foreach (var lyrics in lyricsPerSlide)
        {
            yield return (SlideLayout.Ondertiteling, new Dictionary<string, string>
            {
                { nameof(Liedtekst), lyrics }
            });
        }

        yield return (SlideLayout.Ondertiteling, new Dictionary<string, string>
        {
            { nameof(Liedtekst), string.Empty }
        });
    }
}

public record Collection : PresentationPart
{
    public string EersteDoel
    {
        get => PlaceholderValues.TryGetValue(nameof(EersteDoel), out var value) ? value : string.Empty;
        set => PlaceholderValues[nameof(EersteDoel)] = value;
    }
    public string TweedeDoel
    {
        get => PlaceholderValues.TryGetValue(nameof(TweedeDoel), out var value) ? value : string.Empty;
        set => PlaceholderValues[nameof(TweedeDoel)] = value;
    }

    public override IEnumerable<(SlideLayout Layout, Dictionary<string, string> PlaceholderValues)> GetSlides(bool isPriorToService = false)
    {
        var layout = isPriorToService ? SlideLayout.CollecteTweeDoelenVooraf : SlideLayout.CollecteTweeDoelen;
        
        if (string.IsNullOrEmpty(TweedeDoel))
        {
            layout = isPriorToService ? SlideLayout.CollecteEenDoelVooraf : SlideLayout.CollecteEenDoel;
        }

        yield return (layout, PlaceholderValues);
    }
}

public record Prayer : PresentationPart
{
    public override IEnumerable<(SlideLayout Layout, Dictionary<string, string> PlaceholderValues)> GetSlides(bool isPriorToService = false) => [(SlideLayout.Gebed, PlaceholderValues)];
}

public record BibleReading : PresentationPart
{
    public string Title
    {
        get => PlaceholderValues.TryGetValue(nameof(Title), out var value) ? value : string.Empty;
        set => PlaceholderValues[nameof(Title)] = value;
    }
    public string Text
    {
        get => PlaceholderValues.TryGetValue(nameof(Text), out var value) ? value : string.Empty;
        set => PlaceholderValues[nameof(Text)] = value;
    }

    public override IEnumerable<(SlideLayout Layout, Dictionary<string, string> PlaceholderValues)> GetSlides(bool isPriorToService = false)
    {
        yield return (SlideLayout.BijbellezenAankondiging, PlaceholderValues);
        yield return (SlideLayout.Bijbeltekst, PlaceholderValues);
    }
}

public record TrustAndGreeting : PresentationPart
{
    public new Dictionary<string, string> PlaceholderValues { get; } = new()
    {
        { "Titel", "vertrouwen & groet" }
    };
    
    public override IEnumerable<(SlideLayout Layout, Dictionary<string, string> PlaceholderValues)> GetSlides(bool isPriorToService = false) => [(SlideLayout.PaarsMetTitel, PlaceholderValues)];
}

public record ChildrenMoment : PresentationPart
{
    public string Koffermomenter
    {
        get => PlaceholderValues.TryGetValue(nameof(Koffermomenter), out var value) ? value : string.Empty;
        set => PlaceholderValues[nameof(Koffermomenter)] = value;
    }

    public override IEnumerable<(SlideLayout Layout, Dictionary<string, string> PlaceholderValues)> GetSlides(bool isPriorToService = false) => [(SlideLayout.Koffermoment, PlaceholderValues)];
}

public class PresentationPartConverter : JsonConverter<PresentationPart>
{
    public override PresentationPart Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
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

    public override void Write(Utf8JsonWriter writer, PresentationPart value, JsonSerializerOptions options)
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
    CollecteEenDoel,
    CollecteTweeDoelen,
    CollecteEenDoelVooraf,
    CollecteTweeDoelenVooraf,
    TotZiensMetGebed,
    TotZiens
}