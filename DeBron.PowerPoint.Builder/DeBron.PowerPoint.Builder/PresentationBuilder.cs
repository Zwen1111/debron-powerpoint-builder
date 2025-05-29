using DeBron.PowerPoint.Builder.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using SlideLayout = DeBron.PowerPoint.Builder.Models.SlideLayout;

namespace DeBron.PowerPoint.Builder;

public class PresentationBuilder
{
    private readonly PresentationPart _presentationPart;
    private readonly Dictionary<SlideLayout, (SlideLayoutPart LayoutPart, Slide Slide)> _slidePartsById;
    private readonly string _fileName = $"{Guid.NewGuid()}.pptx";
    private readonly PresentationDocument _presentationDocument;
    private uint _maxSlideId = 256;

    private readonly SlideLayout[] _slideLayoutOrder =
    [
        SlideLayout.WelkomVooraf,
        SlideLayout.LiturgieVooraf,
        SlideLayout.CollecteVooraf,
        SlideLayout.Thema,
        SlideLayout.Welkom,
        SlideLayout.Liturgie,
        SlideLayout.PaarsMetTitel,
        SlideLayout.TussenSlide,
        SlideLayout.LiedAankondiging,
        SlideLayout.LiedAankondigingOverlay,
        SlideLayout.Ondertiteling,
        SlideLayout.Gebed,
        SlideLayout.BlauwMetTitel,
        SlideLayout.LuisterLiedAankondiging,
        SlideLayout.WitMetLiedtekst,
        SlideLayout.Koffermoment,
        SlideLayout.BijbellezenAankondiging,
        SlideLayout.Bijbeltekst,
        SlideLayout.CollecteTweeDoelen,
        SlideLayout.CollecteEenDoel,
        SlideLayout.TotZiensMetGebed,
        SlideLayout.TotZiens
    ];

    public PresentationBuilder()
    {
        File.Copy("template.pptx", _fileName, true);

        _presentationDocument = PresentationDocument.Open(_fileName, true);
        _presentationPart = _presentationDocument.PresentationPart!;

        var slideIdList = _presentationPart.Presentation.SlideIdList;

        _slidePartsById = (slideIdList?.OfType<SlideId>().Select(slideId =>
        {
            var slidePart = (SlidePart)_presentationPart.GetPartById(slideId.RelationshipId!);

            return (slidePart.SlideLayoutPart!, (Slide)slidePart.Slide.CloneNode(true));
        }).ToList() ?? []).Zip(_slideLayoutOrder).ToDictionary(x => x.Second, x => x.First);
        
        RemoveExistingSlides();
    }
    
    public string Build(List<IPresentationPart> parts)
    {
        foreach (var presentationPart in parts)
        {
            switch (presentationPart)
            {
                case Song { UseSubtitle: true } song:
                    AddSongWithSubtitles(song);
                    break;
                case Song song:
                    AddTemplateSlideAndReplaceText(_slidePartsById[SlideLayout.LiedAankondiging], new Dictionary<string, string>
                    {
                        { "Titel", song.Name },
                        { "Ondertitel", string.Empty }
                    });
                    AddTemplateSlideAndReplaceText(_slidePartsById[SlideLayout.WitMetLiedtekst], new Dictionary<string, string>
                    {
                        { "Liedtekst", song.Lyrics }
                    });
                    break;
                case Collection collection:
                    var layout = string.IsNullOrWhiteSpace(collection.SecondGoal)
                            ? SlideLayout.CollecteEenDoel
                            : SlideLayout.CollecteTweeDoelen;
                    
                    AddTemplateSlideAndReplaceText(_slidePartsById[layout], new Dictionary<string, string>
                    {
                        { "Collecte1", collection.FirstGoal },
                        { "Collecte2", collection.SecondGoal }
                    });
                    break;
                case Prayer _:
                    AddTemplateSlideAndReplaceText(_slidePartsById[SlideLayout.Gebed], new Dictionary<string, string>());
                    break;
                case BibleReading bibleReading:
                    AddTemplateSlideAndReplaceText(_slidePartsById[SlideLayout.Bijbeltekst], new Dictionary<string, string>
                    {
                        { "Titel", bibleReading.Title },
                        { "Tekst", bibleReading.Text }
                    });
                    break;
                case TrustAndGreeting _:
                    AddTemplateSlideAndReplaceText(_slidePartsById[SlideLayout.PaarsMetTitel], new Dictionary<string, string>
                    {
                        { "Titel", "vertrouwen & groet" }
                    });
                    break;
                case ChildrenMoment childrenMoment:
                    AddTemplateSlideAndReplaceText(_slidePartsById[SlideLayout.Koffermoment], new Dictionary<string, string>
                    {
                        { "Koffermomenter", childrenMoment.Person }
                    });
                    break;
                default:
                    throw new NotSupportedException($"Unsupported presentation part type: {presentationPart.GetType().Name}");
            }
        }

        _presentationDocument.Dispose();

        return _fileName;
    }

    private void AddSongWithSubtitles(Song song)
    {
        AddTemplateSlideAndReplaceText(_slidePartsById[SlideLayout.LiedAankondigingOverlay], new Dictionary<string, string>
        {
            { "Titel", song.Name },
            { "Ondertitel", song.Subtitle }
        });

        AddTemplateSlideAndReplaceText(_slidePartsById[SlideLayout.Ondertiteling], new Dictionary<string, string>
        {
            { "Liedtekst", string.Empty }
        });

        var lyricsPerSlide = song.Lyrics.Trim().Split("\n\n").SelectMany<string, string>(x => x.StartsWith("\n") ? [string.Empty, x.Trim()] : [x.Trim()]).ToList();

        foreach (var lyrics in lyricsPerSlide)
        {
            AddTemplateSlideAndReplaceText(_slidePartsById[SlideLayout.Ondertiteling], new Dictionary<string, string>
            {
                { "Liedtekst", lyrics }
            });
        }

        AddTemplateSlideAndReplaceText(_slidePartsById[SlideLayout.Ondertiteling], new Dictionary<string, string>
        {
            { "Liedtekst", string.Empty }
        });
    }

    private void RemoveExistingSlides()
    {
        var slideIds = _presentationPart.Presentation.SlideIdList!.ChildElements.OfType<SlideId>().ToList();
        foreach (var slideId in slideIds)
        {
            var slidePart = (SlidePart)_presentationPart.GetPartById(slideId.RelationshipId!);
            _presentationPart.DeletePart(slidePart);
        }

        _presentationPart.Presentation.SlideIdList.RemoveAllChildren();
    }

    private void AddTemplateSlideAndReplaceText(
        (SlideLayoutPart LayoutPart, Slide Slide) slidePart,
        Dictionary<string, string> replacements)
    {
        var newSlidePart = CopySlide(slidePart);

        var texts = newSlidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>().ToList();
        foreach (var (placeholder, value) in replacements)
        {
            foreach (var text in texts)
            {
                text.Text = text.Text.Replace($"{{{{{placeholder}}}}}", value);
            }
        }

        newSlidePart.Slide.Save();
    }

    private SlidePart CopySlide((SlideLayoutPart LayoutPart, Slide Slide) slidePart)
    {
        var newSlidePart = _presentationPart.AddNewPart<SlidePart>();
        newSlidePart.Slide = (Slide)slidePart.Slide.CloneNode(true);

        newSlidePart.AddPart(slidePart.LayoutPart);

        var relId = _presentationPart.GetIdOfPart(newSlidePart);
        var newSlideId = new SlideId
        {
            Id = ++_maxSlideId,
            RelationshipId = relId
        };

        _presentationPart.Presentation.SlideIdList!.Append(newSlideId);

        return newSlidePart;
    }
}
