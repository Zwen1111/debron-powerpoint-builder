using System.Text.Json;
using System.Text.Json.Serialization;

namespace DeBron.PowerPoint.Builder.Models;

public abstract record PresentationPart
{
    public Guid Id { get; set; } = Guid.NewGuid();
    internal readonly Dictionary<string, List<StringReplaceValue>> PlaceholderValues = new();

    public abstract IEnumerable<(SlideLayout Layout, Dictionary<string, List<StringReplaceValue>> PlaceholderValues)> GetSlides(bool isPriorToService = false);
}

public record Song : PresentationPart
{
    public bool UseSubtitle { get; set; } = false;

    public string Titel
    {
        get => PlaceholderValues.TryGetValue(nameof(Titel), out var value) ? value.Single().Value : string.Empty;
        set => PlaceholderValues[nameof(Titel)] = [new StringReplaceValue(value)];
    }

    public string Ondertitel
    {
        get => PlaceholderValues.TryGetValue(nameof(Ondertitel), out var value) ? value.Single().Value : string.Empty;
        set => PlaceholderValues[nameof(Ondertitel)] = [new StringReplaceValue(value)];
    }
    public string Liedtekst
    {
        get => PlaceholderValues.TryGetValue(nameof(Liedtekst), out var value) ? value.Single().Value : string.Empty;
        set => PlaceholderValues[nameof(Liedtekst)] = [new StringReplaceValue(value)];
    }

    public override IEnumerable<(SlideLayout Layout, Dictionary<string, List<StringReplaceValue>> PlaceholderValues)> GetSlides(bool isPriorToService = false)
    {
        yield return (SlideLayout.LiedAankondigingOverlay, PlaceholderValues);

        yield return (SlideLayout.Ondertiteling, new Dictionary<string, List<StringReplaceValue>>
        {
            { nameof(Liedtekst), [new StringReplaceValue(string.Empty)] }
        });

        var lyricsPerSlide = Liedtekst.Trim().Split("\n\n")
            .SelectMany<string, string>(x => x.Trim().StartsWith("\n") ? [string.Empty, x.Trim()] : [x.Trim()]).ToList();

        foreach (var lyrics in lyricsPerSlide)
        {
            yield return (SlideLayout.Ondertiteling, new Dictionary<string, List<StringReplaceValue>>
            {
                { nameof(Liedtekst), [new StringReplaceValue(lyrics)] }
            });
        }

        yield return (SlideLayout.Ondertiteling, new Dictionary<string, List<StringReplaceValue>>
        {
            { nameof(Liedtekst), [new StringReplaceValue(string.Empty)] }
        });
    }
}

public record Collection : PresentationPart
{
    public string EersteDoel
    {
        get => PlaceholderValues.TryGetValue(nameof(EersteDoel), out var value) ? value.Single().Value : string.Empty;
        set => PlaceholderValues[nameof(EersteDoel)] = [new StringReplaceValue(value)];
    }
    public string TweedeDoel
    {
        get => PlaceholderValues.TryGetValue(nameof(TweedeDoel), out var value) ? value.Single().Value : string.Empty;
        set => PlaceholderValues[nameof(TweedeDoel)] = [new StringReplaceValue(value)];
    }

    public override IEnumerable<(SlideLayout Layout, Dictionary<string, List<StringReplaceValue>> PlaceholderValues)> GetSlides(bool isPriorToService = false)
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
    public override IEnumerable<(SlideLayout Layout, Dictionary<string, List<StringReplaceValue>> PlaceholderValues)> GetSlides(bool isPriorToService = false) => [(SlideLayout.Gebed, PlaceholderValues)];
}

public record BibleReading : PresentationPart
{
    public string Reader { get; set; }
    public string BiblebookName { get; set; }
    public int? Chapter { get; set; }
    public int? StartVerse { get; set; }
    public int? EndVerse { get; set; }

    public override IEnumerable<(SlideLayout Layout, Dictionary<string, List<StringReplaceValue>> PlaceholderValues)> GetSlides(bool isPriorToService = false)
    {
        if (!Chapter.HasValue || !StartVerse.HasValue || !EndVerse.HasValue) throw new ArgumentNullException();

        PlaceholderValues["Bijbellezer"] =
            [new StringReplaceValue(Reader)];
        PlaceholderValues["Bijbelgedeelte"] =
            [new StringReplaceValue($"{BiblebookName} {Chapter} : {StartVerse} - {EndVerse}")];
        yield return (SlideLayout.BijbellezenAankondiging, PlaceholderValues);
        
        var verses = BibletextProvider.Provide(Constants.Biblebooks[BiblebookName], Chapter.Value, StartVerse.Value, EndVerse.Value);

        var queue = new Queue<Verse>(verses);

        var versesOnCurrentSlide = new List<StringReplaceValue>();

        var amountOfCharactersOnCurrentSlide = 0;

        while (queue.Any())
        {
            var verse = queue.Dequeue();

            var sentences = verse.Text.Split(". ");

            for (int i = 0; i < sentences.Length; i++)
            {
                var sentence = sentences[i];
                
                var textLength =  sentence.Length + (i == 0 ? verse.Number.ToString().Length : 0);
                
                if (amountOfCharactersOnCurrentSlide + textLength <= 425)
                {
                    if (i == 0)
                    {
                        versesOnCurrentSlide.Add(new StringReplaceValue(verse.Number.ToString(), true));
                    }

                    versesOnCurrentSlide.Add(new StringReplaceValue(sentence));
                    
                    amountOfCharactersOnCurrentSlide += textLength;
                }
                else
                {
                    PlaceholderValues["Bijbeltekst"] = versesOnCurrentSlide;
                    yield return (SlideLayout.Bijbeltekst, PlaceholderValues);
                    
                    versesOnCurrentSlide = [];
                    amountOfCharactersOnCurrentSlide = 0;
                }
            }
        }

        if (versesOnCurrentSlide.Any())
        {
            PlaceholderValues["Bijbeltekst"] = versesOnCurrentSlide;
            yield return (SlideLayout.Bijbeltekst, PlaceholderValues);
        }
    }
}

public record TrustAndGreeting : PresentationPart
{
    public new Dictionary<string, List<StringReplaceValue>> PlaceholderValues { get; } = new()
    {
        { "Titel", [new StringReplaceValue("vertrouwen & groet")] }
    };
    
    public override IEnumerable<(SlideLayout Layout, Dictionary<string, List<StringReplaceValue>> PlaceholderValues)> GetSlides(bool isPriorToService = false) => [(SlideLayout.PaarsMetTitel, PlaceholderValues)];
}

public record ChildrenMoment : PresentationPart
{
    public string Koffermomenter
    {
        get => PlaceholderValues.TryGetValue(nameof(Koffermomenter), out var value) ? value.Single().Value : string.Empty;
        set => PlaceholderValues[nameof(Koffermomenter)] = [new StringReplaceValue(value)];
    }

    public override IEnumerable<(SlideLayout Layout, Dictionary<string, List<StringReplaceValue>> PlaceholderValues)> GetSlides(bool isPriorToService = false) => [(SlideLayout.Koffermoment, PlaceholderValues)];
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

public record StringReplaceValue(string Value, bool Superscript = false);