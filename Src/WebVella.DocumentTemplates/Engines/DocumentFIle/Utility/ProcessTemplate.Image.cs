using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Word = DocumentFormat.OpenXml.Wordprocessing;


namespace WebVella.DocumentTemplates.Engines.DocumentFile.Utility;

public partial class WvDocumentFileEngineUtility
{
    public void _copyImages(WordprocessingDocument original, WordprocessingDocument target)
    {
        var originalMainPart = original.MainDocumentPart;
        var targetMainPart = target.MainDocumentPart;

        if (originalMainPart == null || targetMainPart == null)
            throw new InvalidOperationException("Both documents must have a MainDocumentPart.");
        Dictionary<string, string> imageIdMap = new Dictionary<string, string>();
        _copyImageParts(originalMainPart, targetMainPart, imageIdMap);
        _copyImagesFromHeaderFooter(originalMainPart, targetMainPart, imageIdMap);
        _updateBlipReferences(targetMainPart, imageIdMap);
    }

    private void _copyImageParts(MainDocumentPart originalMainPart, MainDocumentPart targetMainPart,
        Dictionary<string, string> imageIdMap)
    {
        foreach (var imagePart in originalMainPart.ImageParts)
        {
            string originalImageId = originalMainPart.GetIdOfPart(imagePart);
            ImagePart newImagePart = targetMainPart.AddImagePart(imagePart.ContentType);
            var inputStream = imagePart.GetStream();
            var outputStream = newImagePart.GetStream();
            _copyStream(inputStream, outputStream);
            inputStream.Close();
            outputStream.Close();
            string newImageId = targetMainPart.GetIdOfPart(newImagePart);
            imageIdMap[originalImageId] = newImageId;
        }
    }

    private void _copyImagesFromHeaderFooter(MainDocumentPart originalMainPart, MainDocumentPart targetMainPart,
        Dictionary<string, string> imageIdMap)
    {
        var originalHeaderParts = originalMainPart.HeaderParts.ToList();
        var targetHeaderParts = targetMainPart.HeaderParts.ToList();
        foreach (var originalPart in originalHeaderParts)
        {
            var targetPart = targetHeaderParts.FirstOrDefault(x => x.Uri == originalPart.Uri);
            if (targetPart is null) continue;
            _copyHeaderImages(originalPart, targetPart, imageIdMap);

        }

        var originalFooterParts = originalMainPart.FooterParts.ToList();
        var targetFooterParts = targetMainPart.FooterParts.ToList();
        foreach (var originalPart in originalFooterParts)
        {
            var targetPart = targetFooterParts.FirstOrDefault(x => x.Uri == originalPart.Uri);
            if (targetPart is null) continue;
            _copyFooterImages(originalPart, targetPart, imageIdMap);

        }
    }

    private void _copyHeaderImages(HeaderPart originalPart, HeaderPart targetPart,
        Dictionary<string, string> imageIdMap)
    {
        targetPart.DeleteParts<ImagePart>(targetPart.ImageParts);
        foreach (var originalHeader in originalPart.Parts) {
            // 'rel.OpenXmlPart' is the actual part (e.g., ImagePart)
            // 'rel.RelationshipId' is the ID used in the source XML (e.g., "rId1")
            if (originalHeader.OpenXmlPart is ImagePart sourceImagePart)
            {
                // Create a matching ImagePart in the destination header
                ImagePart newImagePart = targetPart.AddImagePart(sourceImagePart.ContentType);

                // Copy the binary data (stream)
                using (Stream stream = sourceImagePart.GetStream())
                {
                    newImagePart.FeedData(stream);
                }

                // Get the NEW Relationship ID generated in the destination
                string newRelId = targetPart.GetIdOfPart(newImagePart);

                // Update the cloned XML to use the new ID instead of the old one
                // This searches the XML for the old ID and swaps it
                UpdateRelationshipId(targetPart.Header, originalHeader.RelationshipId, newRelId);
                imageIdMap[originalHeader.RelationshipId] = newRelId;
            }
        }
    }

    private void _copyFooterImages(FooterPart originalPart, FooterPart targetPart,
        Dictionary<string, string> imageIdMap)
    {
        targetPart.DeleteParts<ImagePart>(targetPart.ImageParts);
        foreach (var originalFooter in originalPart.Parts)
        {
            // 'rel.OpenXmlPart' is the actual part (e.g., ImagePart)
            // 'rel.RelationshipId' is the ID used in the source XML (e.g., "rId1")
            if (originalFooter.OpenXmlPart is ImagePart sourceImagePart)
            {
                // Create a matching ImagePart in the destination header
                ImagePart newImagePart = targetPart.AddImagePart(sourceImagePart.ContentType);

                // Copy the binary data (stream)
                using (Stream stream = sourceImagePart.GetStream())
                {
                    newImagePart.FeedData(stream);
                }

                // Get the NEW Relationship ID generated in the destination
                string newRelId = targetPart.GetIdOfPart(newImagePart);

                // Update the cloned XML to use the new ID instead of the old one
                // This searches the XML for the old ID and swaps it
                UpdateRelationshipId(targetPart.Footer, originalFooter.RelationshipId, newRelId);
                imageIdMap[originalFooter.RelationshipId] = newRelId;
            }
        }
    }

    private void _updateBlipReferences(MainDocumentPart targetMainPart, Dictionary<string, string> imageIdMap)
    {
        foreach (var targetHeader in targetMainPart.HeaderParts)
        {
            foreach (var drawing in targetHeader.Header.Descendants<Word.Drawing>())
            {
                var blip = drawing.Descendants<Blip>().FirstOrDefault();
                if (blip != null && blip.Embed != null)
                {
                    if (imageIdMap.ContainsKey(blip.Embed.Value))
                    {
                        blip.Embed.Value = imageIdMap[blip.Embed.Value];
                    }
                    else
                    {

                    }
                }
            }
        }

        foreach (var targetFooter in targetMainPart.FooterParts)
        {
            foreach (var drawing in targetFooter.Footer.Descendants<Word.Drawing>())
            {
                var blip = drawing.Descendants<Blip>().FirstOrDefault();
                if (blip != null && blip.Embed != null && imageIdMap.ContainsKey(blip.Embed.Value))
                {
                    blip.Embed.Value = imageIdMap[blip.Embed.Value];
                }
            }
        }

        foreach (var drawing in targetMainPart.Document.Body.Descendants<Word.Drawing>())
        {
            var blip = drawing.Descendants<Blip>().FirstOrDefault();
            if (blip != null && blip.Embed != null && imageIdMap.ContainsKey(blip.Embed.Value))
            {
                blip.Embed.Value = imageIdMap[blip.Embed.Value];
            }
        }
    }

    private void _copyStream(Stream input, Stream output)
    {
        input.CopyTo(output);
        output.Flush();
    }
    // Helper method to traverse XML and replace Relationship IDs
    private static void UpdateRelationshipId(OpenXmlElement root, string oldId, string newId)
    {
        // Search for elements that use relationships (Blip for images, etc.)
        // Note: We search for attributes named "embed" or "id" in specific namespaces

        foreach (var element in root.Descendants())
        {
            foreach (var attr in element.GetAttributes())
            {
                if (attr.Value == oldId)
                {
                    // Common attributes for relationships are r:embed, r:id, or r:link
                    if (attr.LocalName == "embed" || attr.LocalName == "id" || attr.LocalName == "link")
                    {
                        element.SetAttribute(new OpenXmlAttribute(attr.Prefix, attr.LocalName, attr.NamespaceUri, newId));
                    }
                }
            }
        }
    }
}