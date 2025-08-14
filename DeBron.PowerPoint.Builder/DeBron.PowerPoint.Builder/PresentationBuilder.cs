using System.Text.RegularExpressions;
using DeBron.PowerPoint.Builder.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Paragraph = DocumentFormat.OpenXml.Drawing.Paragraph;
using PresentationPart = DeBron.PowerPoint.Builder.Models.PresentationPart;
using Run = DocumentFormat.OpenXml.Drawing.Run;
using SlideLayout = DeBron.PowerPoint.Builder.Models.SlideLayout;
using Text = DocumentFormat.OpenXml.Drawing.Text;
using VerticalTextAlignment = DocumentFormat.OpenXml.Spreadsheet.VerticalTextAlignment;

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
        Dictionary<string, List<StringReplaceValue>> replacements)
    {
        var newSlidePart = CopySlide(slidePart);

        foreach (var (placeholder, values) in replacements)
        {
            var placeholderPattern = new Regex($"{{{{{placeholder}}}}}", RegexOptions.IgnoreCase);
            
            var paragraphs = newSlidePart.Slide.Descendants<Paragraph>().ToList();

            foreach (var paragraph in paragraphs)
            {
                var runs = paragraph.Descendants<Run>().ToList();

                var newRuns = new List<Run>();

                foreach (var run in runs)
                {
                    var matches = placeholderPattern.Matches(run.Text.Text).ToList();

                    if (matches.Count == 0)
                    {
                        newRuns.Add(run);
                        continue;
                    }

                    foreach (var match in matches)
                    {
                        var index = match.Index;
                        var length = match.Length;

                        var preRun = run.CloneNode(true) as Run;
                        preRun.RemoveAllChildren<Text>();
                        var preRunText = new Text(run.Text.Text.Substring(0, index));

                        preRun.AppendChild(preRunText);
                        newRuns.Add(preRun);

                        newRuns.AddRange(values.Select(replacementValue =>
                        {
                            var clonedRun = run.CloneNode(true) as Run;
                            clonedRun.RemoveAllChildren<Text>();

                            if (replacementValue.Superscript)
                            {
                                clonedRun.RunProperties.Baseline = clonedRun.RunProperties.FontSize * 12;
                            }

                            var clonedText = new Text(replacementValue.Value.ToString());

                            clonedRun.AppendChild(clonedText);

                            return clonedRun;
                        }));

                        var postRun = run.CloneNode(true) as Run;
                        postRun.RemoveAllChildren<Text>();

                        var postRunText = new Text(run.Text.Text.Substring(index + length));

                        postRun.AppendChild(postRunText);

                        newRuns.Add(postRun);
                    }
                }
                
                paragraph.RemoveAllChildren<Run>();
                
                foreach (var newRun in newRuns.Where(r => r.Text.Text.Length > 0))
                {
                    paragraph.AppendChild(newRun);
                }
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
