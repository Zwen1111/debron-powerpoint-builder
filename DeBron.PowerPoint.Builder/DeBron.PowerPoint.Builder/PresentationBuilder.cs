using DeBron.PowerPoint.Builder.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace DeBron.PowerPoint.Builder;

public class PresentationBuilder
{
    private readonly PresentationPart _presentationPart;
    private readonly List<(SlideLayoutPart LayoutPart, Slide Slide)> _slideParts;
    private readonly string _fileName = $"{Guid.NewGuid()}.pptx";
    private readonly PresentationDocument _presentationDocument;
    private uint _maxSlideId = 256;

    public PresentationBuilder()
    {
        File.Copy("template.pptx", _fileName, true);

        _presentationDocument = PresentationDocument.Open(_fileName, true);
        _presentationPart = _presentationDocument.PresentationPart!;

        _slideParts = _presentationPart.SlideParts.Select(s => (s.SlideLayoutPart!, (Slide)s.Slide.CloneNode(true))).ToList();
        
        RemoveExistingSlides();
    }
    
    public string Build(List<Song> songs)
    {
        songs.ForEach(AddSongWithSubtitles);

        _presentationDocument.Dispose();

        return _fileName;
    }

    private void AddSongWithSubtitles(Song song)
    {
        AddTemplateSlideAndReplaceText(_slideParts[1], new Dictionary<string, string>
        {
            { "Titel", song.Name },
            { "Ondertitel", song.Subtitle }
        });

        AddTemplateSlideAndReplaceText(_slideParts[0], new Dictionary<string, string>
        {
            { "Liedtekst", string.Empty }
        });

        var lyricsPerSlide = song.Lyrics.Split("\n\n");

        foreach (var lyrics in lyricsPerSlide)
        {
            AddTemplateSlideAndReplaceText(_slideParts[0], new Dictionary<string, string>
            {
                { "Liedtekst", lyrics }
            });
        }

        AddTemplateSlideAndReplaceText(_slideParts[0], new Dictionary<string, string>
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
