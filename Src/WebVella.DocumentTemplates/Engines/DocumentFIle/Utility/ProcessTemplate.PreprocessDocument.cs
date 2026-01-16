using System.Data;
using System.Globalization;
using WebVella.DocumentTemplates.Engines.Text;
using Word = DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;

namespace WebVella.DocumentTemplates.Engines.DocumentFile.Utility;

public partial class WvDocumentFileEngineUtility
{
    public static void IsolateTemplateTags(WordprocessingDocument doc)
    {
        var body = doc.MainDocumentPart.Document.Body;

        foreach (var paragraph in body.Descendants<Word.Paragraph>())
        {
            ProcessParagraph(paragraph);
        }

        doc.MainDocumentPart.Document.Save();
    }

    private static void ProcessParagraph(Word.Paragraph paragraph)
    {
        // 1. Build the character map
        var textMap = new List<(char Character, Word.Text TextElement, int IndexInElement)>();
        foreach (var textElem in paragraph.Descendants<Word.Text>())
        {
            string textContent = textElem.Text;
            for (int i = 0; i < textContent.Length; i++)
            {
                textMap.Add((textContent[i], textElem, i));
            }
        }

        string fullText = new string(textMap.Select(m => m.Character).ToArray());

        // 2. Find matches
        var matches = Regex.Matches(fullText, @"\{\{.*?\}\}");
        if (matches.Count == 0) return;

        // 3. Global tracker for consumed indices to prevent duplication
        // Any character index marked true here will be removed from its original node.
        bool[] globalConsumed = new bool[fullText.Length];

        // Process matches in DESCENDING order (Right-to-Left).
        // This is critical so we don't mess up insertion points for earlier parts of the text.
        foreach (Match match in matches.Cast<Match>().OrderByDescending(m => m.Index))
        {
            int startIndex = match.Index;
            int length = match.Length;
            int endIndex = startIndex + length - 1;

            // Identify Start and End nodes from the map
            var startNodeInfo = textMap[startIndex];
            var endNodeInfo = textMap[endIndex];

            Word.Text startTextElem = startNodeInfo.TextElement;
            Word.Text endTextElem = endNodeInfo.TextElement;
            Word.Run startParentRun = (Word.Run)startTextElem.Parent;
            Word.Run endParentRun = (Word.Run)endTextElem.Parent;

            // --- A. Handle Suffix (Text occurring AFTER the tag in the same node) ---
            // If the tag ends in the middle of a node, the remaining text needs to move to a new Run
            // to stay *after* our new tag.
            // Check if there is valid text after this match in the End Node.

            // We look at the map indices belonging to the End Node that are > endIndex
            var suffixIndices = textMap
                .Select((x, idx) => new { Info = x, GlobalIndex = idx })
                .Where(x => x.Info.TextElement == endTextElem && x.GlobalIndex > endIndex)
                .ToList();

            // Only create a suffix run if there are characters that are NOT already consumed by a previous match
            // (Remember: we iterate backwards, so 'previous' operations processed text to the right)
            string suffixText = "";
            bool hasValidSuffix = false;
            foreach (var idx in suffixIndices)
            {
                if (!globalConsumed[idx.GlobalIndex])
                {
                    suffixText += idx.Info.Character;
                    hasValidSuffix = true;
                    // Mark these as consumed from the original node so they don't stay there
                    globalConsumed[idx.GlobalIndex] = true;
                }
            }

            if (hasValidSuffix && !string.IsNullOrEmpty(suffixText))
            {
                Word.Run suffixRun = new Word.Run();
                // Copy properties from the end run
                if (endParentRun.RunProperties != null)
                    suffixRun.AppendChild(endParentRun.RunProperties.CloneNode(true));

                Word.Text suffixTextElem = new Word.Text(suffixText) { Space = SpaceProcessingModeValues.Preserve };
                suffixRun.AppendChild(suffixTextElem);

                // Insert Suffix AFTER the end run. 
                // Note: If we are splitting one run into 3 (Prefix-Tag-Suffix), 
                // this suffix goes at the very end.
                endParentRun.Parent.InsertAfter(suffixRun, endParentRun);
            }

            // --- B. Create and Insert the Tag Run ---
            Word.Run tagRun = new Word.Run();
            // Copy properties from the start run (Leading {{)
            if (startParentRun.RunProperties != null)
                tagRun.AppendChild(startParentRun.RunProperties.CloneNode(true));

            Word.Text tagTextElem = new Word.Text(match.Value) { Space = SpaceProcessingModeValues.Preserve };
            tagRun.AppendChild(tagTextElem);

            // Insert the tag AFTER the start run.
            // Logic: The Start Run will eventually be trimmed to contain only "Prefix".
            // So "Prefix" -> "Tag" -> "Suffix".
            startParentRun.Parent.InsertAfter(tagRun, startParentRun);

            // --- C. Mark Tag Indices as Consumed ---
            for (int i = startIndex; i <= endIndex; i++)
            {
                globalConsumed[i] = true;
            }
        }

        // --- D. Cleanup Phase ---
        // Now we simply visit every Text node touched by ANY match and rebuild its text
        // excluding any character marked as 'globalConsumed'.

        var allTouchedNodes = textMap
            .Where((x, i) => globalConsumed[i]) // Only care about nodes that have consumed chars
            .Select(x => x.TextElement)
            .Distinct()
            .ToList();

        foreach (var node in allTouchedNodes)
        {
            // Find all global indices belonging to this node
            var nodeIndices = textMap
                .Select((x, idx) => new { Info = x, GlobalIndex = idx })
                .Where(x => x.Info.TextElement == node)
                .ToList();

            string newContent = "";
            foreach (var item in nodeIndices)
            {
                // If NOT consumed, keep it.
                if (!globalConsumed[item.GlobalIndex])
                {
                    newContent += item.Info.Character;
                }
            }

            node.Text = newContent;
        }

        // --- E. Final Polish ---
        // Remove empty runs that might be left over (e.g. if a run contained ONLY the tag)
        foreach (var node in allTouchedNodes)
        {
            if (string.IsNullOrEmpty(node.Text))
            {
                // If the Text element is empty, remove it.
                // If the parent Run becomes empty, remove that too.
                var parent = node.Parent;
                node.Remove();

                if (parent != null && !parent.HasChildren)
                {
                    parent.Remove();
                }
            }
        }
    }
}