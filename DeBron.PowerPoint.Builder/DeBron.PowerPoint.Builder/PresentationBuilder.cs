using DeBron.PowerPoint.Builder.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Text = DocumentFormat.OpenXml.Presentation.Text;

namespace DeBron.PowerPoint.Builder;

public static class PresentationBuilder
{
    public static void Build(IEnumerable<Song> songs)
    {
        File.Copy("template.pptx", "presentation.pptx", true);

        using var presentationDocument = PresentationDocument.Open("presentation.pptx", true);
        var presentationPart = presentationDocument.PresentationPart!;
        var presentation = presentationPart.Presentation;

        // Haal de eerste slide als sjabloon
        var sourceSlidePart = presentationPart.SlideParts.First();
        var layoutPart = sourceSlidePart.SlideLayoutPart!;

        // Clone de slide XML naar geheugen
        var slideXml = (Slide)sourceSlidePart.Slide.CloneNode(true);

        // Verwijder alle bestaande slides
        var slideIds = presentation.SlideIdList!.ChildElements.OfType<SlideId>().ToList();
        foreach (var slideId in slideIds)
        {
            var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!);
            presentationPart.DeletePart(slidePart);
        }
        presentation.SlideIdList.RemoveAllChildren();

        uint maxSlideId = 256;

        foreach (var song in songs)
        {
            var lyricsPerSlide = song.Lyrics.Split("\n\n");

            foreach (var lyrics in lyricsPerSlide)
            {
                AddTemplateSlideAndReplaceText(presentationPart, slideXml, layoutPart, new Dictionary<string, string>
                {
                    { "ondertiteling", lyrics }
                }, ref maxSlideId);
            }
        }

        presentation.Save();
    }

    private static void AddTemplateSlideAndReplaceText(
        PresentationPart presentationPart,
        Slide slideTemplate,
        SlideLayoutPart layoutPart,
        Dictionary<string, string> replacements,
        ref uint maxSlideId)
    {
        var newSlidePart = CopySlide(presentationPart, slideTemplate, layoutPart, ref maxSlideId);

        var texts = newSlidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>();
        foreach (var (placeholder, value) in replacements)
        {
            foreach (var text in texts)
            {
                text.Text = text.Text.Replace($"{{{{{placeholder}}}}}", value);
            }
        }

        newSlidePart.Slide.Save();
    }

    private static SlidePart CopySlide(
        PresentationPart presentationPart,
        Slide slideTemplate,
        SlideLayoutPart layoutPart,
        ref uint maxSlideId)
    {
        var newSlidePart = presentationPart.AddNewPart<SlidePart>();
        newSlidePart.Slide = (Slide)slideTemplate.CloneNode(true);

        newSlidePart.AddPart(layoutPart);

        var relId = presentationPart.GetIdOfPart(newSlidePart);
        var newSlideId = new SlideId
        {
            Id = ++maxSlideId,
            RelationshipId = relId
        };

        presentationPart.Presentation.SlideIdList!.Append(newSlideId);

        return newSlidePart;
    }

}
