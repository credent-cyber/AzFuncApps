﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PIFunc.DocxHelper
{
    public class EditDocumentHeader
    {
        public static void ModifyHeaderSection(WordprocessingDocument doc, string docid, string ProcedureRef, string RevisionNo, string RevisionDate, string FileName)
        {
            try
            {
                // Remove content controls before updating the document
                RemoveContentControls(doc);

                MainDocumentPart mainPart = doc.MainDocumentPart;

                UpdateHeaderValues(mainPart.HeaderParts, docid, ProcedureRef, RevisionNo, RevisionDate, FileName);

                UpdateFontProperties(mainPart.HeaderParts);// update font and font size

                doc.Save();

            }
            catch (Exception ex)
            {

            }
        }

        private static async void UpdateHeaderValues(IEnumerable<HeaderPart> headerParts, string docid, string ProcedureRef, string RevisionNo, string RevisionDate, string FileName)
        {
            foreach (var headerPart in headerParts)
            {
                foreach (TableRow row in headerPart.Header.Descendants<TableRow>())
                {
                    var cellValues = row.Descendants<TableCell>()
                        .Select(cell => string.Join("", cell.Descendants<Text>().Select(text => text.Text)))
                        .ToList();

                    // Update specific header cell values
                    await UpdateCellValue(row, cellValues, "Doc ID NO:", docid);
                    await UpdateCellValue(row, cellValues, "PROCEDURE REF NO:", ProcedureRef);
                    await UpdateCellValue(row, cellValues, "REVISION NO", RevisionNo);
                    await UpdateCellValue(row, cellValues, "REVISION DATE:", RevisionDate);
                    await UpdateCellValue(row, cellValues, "Document Title", FileName);
                }
            }
        }

        private static async Task UpdateCellValue(TableRow row, List<string> cellValues, string label, string newValue)
        {
            int labelIndex = cellValues.FindIndex(value => value.Contains(label, StringComparison.InvariantCultureIgnoreCase));
            int maxLineLength = 20;
            if (labelIndex != -1 && labelIndex < cellValues.Count - 1)
            {
                var nextCell = row.Descendants<TableCell>().ElementAt(labelIndex + 1);
                var textElement = nextCell.Descendants<Text>().FirstOrDefault();

                if (textElement == null)
                {
                    // Create a new Text element if it doesn't exist
                    textElement = new Text();
                    // Create a new Run element to contain the Text element
                    var run = new Run(textElement);
                    // Add the Run to the Paragraph (or create a new Paragraph if none exists)
                    var paragraph = nextCell.Descendants<Paragraph>().FirstOrDefault();
                    if (paragraph == null)
                    {
                        paragraph = new Paragraph();
                        nextCell.Append(paragraph);
                    }
                    paragraph.Append(run);
                }

                textElement.Space = SpaceProcessingModeValues.Preserve;

                // Wrap the text if the length is more than maxLineLength
                if (newValue.Length > maxLineLength)
                {
                    string wrappedText = WrapText(newValue, maxLineLength);
                    textElement.Text = wrappedText;
                }
                else
                {
                    textElement.Text = newValue;
                }
            }
        }


        private static string WrapText(string text, int maxLineLength)
        {
            StringBuilder wrappedText = new StringBuilder();
            int currentIndex = 0;

            while (currentIndex < text.Length)
            {
                // Determine the remaining length and the max length of the current line
                int remainingLength = text.Length - currentIndex;
                int lineLength = Math.Min(maxLineLength, remainingLength);

                // Find the last space within the max line length
                int wrapAt = text.LastIndexOf(' ', currentIndex + lineLength - 1, lineLength);

                if (wrapAt == -1 || wrapAt < currentIndex)
                {
                    // No space found within the limit, wrap at max length
                    wrapAt = currentIndex + lineLength;
                }
                else
                {
                    // Wrap at the last space within the limit
                    wrapAt = wrapAt + 1; // Include the space character itself
                }

                wrappedText.AppendLine(text.Substring(currentIndex, wrapAt - currentIndex).Trim());

                currentIndex = wrapAt;
            }

            return wrappedText.ToString().TrimEnd(); // Remove the trailing newline
        }


        // Function to remove content control placeholders
        static void RemoveContentControls(WordprocessingDocument doc)
        {

            MainDocumentPart mainPart = doc.MainDocumentPart;

            // Remove content controls in headers
            foreach (var headerPart in mainPart.HeaderParts)
            {
                RemoveContentControls(headerPart.RootElement.Descendants<SdtElement>());
            }

            // Remove content controls in the main document body
            // RemoveContentControls(mainPart.Document.Body.Descendants<SdtElement>());

            // Remove content controls in footers if needed
            foreach (var footerPart in mainPart.FooterParts)
            {
                RemoveContentControls(footerPart.RootElement.Descendants<SdtElement>());
            }

            mainPart.Document.Save();

        }

        static void RemoveContentControls(IEnumerable<SdtElement> contentControls)
        {
            foreach (var contentControl in contentControls)
            {
                // Clear content control properties to remove it
                contentControl.SdtProperties.RemoveAllChildren();
            }
        }

        private static void UpdateFontProperties(IEnumerable<HeaderPart> headerParts)
        {
            foreach (var headerPart in headerParts)
            {
                var header = headerPart.Header;

                foreach (var paragraph in header.Descendants<Paragraph>())
                {
                    foreach (var run in paragraph.Descendants<Run>())
                    {
                        var runProperties = run.RunProperties ?? new RunProperties();

                        // Set font name to Times New Roman and font size to 12
                        runProperties.RunFonts = new RunFonts() { Ascii = "Times New Roman" };
                        runProperties.FontSize = new FontSize() { Val = "18" }; // Font size in half-point measurement, 12 * 2                                                                                
                        runProperties.Color = new Color() { Val = "000000" };// Set font color to black

                        run.RunProperties = runProperties;
                    }
                }
            }
        }


    }
}
