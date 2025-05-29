using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PresentationPart = DeBron.PowerPoint.Builder.Models.PresentationPart;
using SlideLayout = DeBron.PowerPoint.Builder.Models.SlideLayout;

namespace DeBron.PowerPoint.Builder;

public class PresentationBuilder
{
    private readonly DocumentFormat.OpenXml.Packaging.PresentationPart _presentationPart;
    private readonly Dictionary<SlideLayout, (SlideLayoutPart LayoutPart, Slide Slide)> _slidePartsById;
    private readonly string _fileName = $"{Guid.NewGuid()}.pptx";
    private readonly PresentationDocument _presentationDocument;
    private uint _maxSlideId = 256;

    private readonly SlideLayout[] _slideLayoutOrder =
    [
        SlideLayout.WelkomVooraf,
        SlideLayout.LiturgieVooraf,
        SlideLayout.CollecteEenDoelVooraf,
        SlideLayout.CollecteTweeDoelenVooraf,
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
        SlideLayout.CollecteEenDoel,
        SlideLayout.CollecteTweeDoelen,
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
    
    public string Build(List<PresentationPart> parts)
    {
        foreach (var presentationPart in parts)
        {
            var slides = presentationPart.GetSlides();

            foreach (var (slideLayout, placeholderValues) in slides)
            {
                AddTemplateSlideAndReplaceText(_slidePartsById[slideLayout], placeholderValues);
            }
        }

        _presentationDocument.Dispose();

        return _fileName;
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
        var combinedText = string.Concat(texts.Select(t => t.Text));

        foreach (var (placeholder, value) in replacements)
        {
            var placeholderPattern = new Regex($"{{{{{placeholder}}}}}", RegexOptions.IgnoreCase);
            var matches = placeholderPattern.Matches(combinedText);

            var pos = 0;
            var offset = 0;
            foreach (var text in texts)
            {
                var textLength = text.Text.Length;
                
                var alreadyReplaceMatches = matches.Where(m => m.Index < pos && m.Index + m.Length > pos).ToList();
                
                foreach (var alreadyReplaceMatch in alreadyReplaceMatches)
                {
                    var lengthToRemove = Math.Min(alreadyReplaceMatch.Length - pos + alreadyReplaceMatch.Index, text.Text.Length);
                    var textToRemove = text.Text.Substring(0, lengthToRemove);
                    text.Text = text.Text.Replace(textToRemove, string.Empty);
                }
                
                var matchHits = matches.Where(m => m.Index >= pos && m.Index < pos + textLength).ToList();
                
                foreach (var matchHit in matchHits)
                {
                    var index = matchHit.Index + offset - pos;
                    var length = Math.Min(matchHit.Length, text.Text.Length - index);
                    
                    var textToReplace = text.Text.Substring(index, length);
                    offset += value.Length - textToReplace.Length;
                    text.Text = text.Text.Replace(textToReplace, value);
                }

                pos += textLength;
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
