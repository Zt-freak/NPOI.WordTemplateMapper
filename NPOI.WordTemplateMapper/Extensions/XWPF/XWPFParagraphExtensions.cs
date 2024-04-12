using NPOI.XWPF.UserModel;
using System.Text.RegularExpressions;

namespace NPOI.WordTemplateMapper.Extensions.XWPF;

public static class XWPFParagraphExtensions
{
    public static void ReplaceTextKeepLineEndings(this XWPFParagraph xwpfParagraph, string textToReplace, string replacementText)
    {
        string paragraphText = xwpfParagraph.Text;

        if (!paragraphText.Contains(textToReplace))
            return;

        if (!replacementText.Contains('\n'))
        {
            string textToInsert = paragraphText.Replace(textToReplace, replacementText);
            if (!string.IsNullOrEmpty(xwpfParagraph.Text))
                xwpfParagraph.ReplaceText(xwpfParagraph.Text, textToInsert);
            return;
        }

        List<int> positionsOfOldTextOccurences = new();
        positionsOfOldTextOccurences
            .AddRange(
                Regex.Matches(paragraphText, Regex.Escape(textToReplace))
                    .Select(match => match.Index)
            );

        for (int i = positionsOfOldTextOccurences.Count - 1; i >= 0; i--)
        {
            int currentPosition = positionsOfOldTextOccurences[i];
            paragraphText = paragraphText
                .Remove(currentPosition, textToReplace.Length)
                .Insert(currentPosition, replacementText);
        }

        string[] newParagraphsToInsert = paragraphText.Split('\n');
        xwpfParagraph.ReplaceText(xwpfParagraph.Text, string.Empty);

        XWPFRun run = xwpfParagraph.CreateRun();
        foreach (string newParagraph in newParagraphsToInsert)
        {
            run.AppendText(newParagraph);
            run.AddCarriageReturn();
        }
    }
}
