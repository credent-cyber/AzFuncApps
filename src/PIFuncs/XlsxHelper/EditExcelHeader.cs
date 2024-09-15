using ClosedXML.Excel;
using System;
using System.Text;

namespace PIFunc.XlsxHelper
{
    public class EditExcelHeader
    {
        public static void ModifyHeaderSection(XLWorkbook workbook, string docId, string procedureRef, string revisionNo, string revisionDate, string fileName)
        {
            try
            {
                // Loop through all the worksheets in the workbook
                foreach (IXLWorksheet worksheet in workbook.Worksheets)
                {
                    // Iterate over rows 1 to 7 to find and modify the specific headers
                    foreach (IXLRow row in worksheet.Rows(1, 7))
                    {
                        foreach (IXLCell cell in row.Cells())
                        {
                            string cellValue = cell.GetString();

                            // Check for specific keywords and update the adjacent cells
                            if (cellValue.Contains("DOC ID"))
                            {
                                UpdateAdjacentCell(worksheet, cell, docId, false);
                            }
                            else if (cellValue.Contains("PROCEDURE REF"))
                            {
                                UpdateAdjacentCell(worksheet, cell, procedureRef, false);
                            }
                            else if (cellValue.Contains("REVISION NO"))
                            {
                                UpdateAdjacentCell(worksheet, cell, revisionNo, false);
                            }
                            else if (cellValue.Contains("REVISION DATE"))
                            {
                                UpdateAdjacentCell(worksheet, cell, revisionDate, false);
                            }
                            else if (cellValue.Contains("Document Name"))
                            {
                                UpdateAdjacentCell(worksheet, cell, fileName, true); // Bold the Document Name
                            }
                        }
                    }
                }

                // Save the workbook after modification
                workbook.Save();
            }
            catch (Exception ex)
            {
                // Handle exceptions
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to update the adjacent cell with the new value, wrap text, and apply formatting
        private static void UpdateAdjacentCell(IXLWorksheet worksheet, IXLCell cell, string newValue, bool isBold)
        {
            // Get the column and row index of the current cell
            int columnNumber = cell.Address.ColumnNumber;
            int rowNumber = cell.Address.RowNumber;

            // Get the adjacent cell (next column in the same row)
            IXLCell nextCell = worksheet.Cell(rowNumber, columnNumber + 1);

            // Wrap the text before setting it in the cell
            string wrappedText = WrapText(newValue, 50); // Assume a max of 50 characters per line

            // Set the wrapped text in the adjacent cell
            nextCell.Value = wrappedText;

            // Apply the "Times New Roman" font and size 12 to the adjacent cell
            nextCell.Style.Font.FontName = "Times New Roman";
            nextCell.Style.Font.FontSize = 12;

            // Apply bold formatting if required (for "Document Name")
            if (isBold)
            {
                nextCell.Style.Font.Bold = true;
            }
        }

        // Function to wrap text based on target line length
        private static string WrapText(string text, int targetLineLength)
        {
            StringBuilder wrappedText = new StringBuilder();
            int currentIndex = 0;

            while (currentIndex < text.Length)
            {
                // Determine the remaining length of the text
                int remainingLength = text.Length - currentIndex;
                // If remaining text is less than the target length, just append the remaining text
                if (remainingLength <= targetLineLength)
                {
                    wrappedText.AppendLine(text.Substring(currentIndex).Trim());
                    break;
                }

                // Try to find the next space within the target range or slightly beyond it
                int wrapAt = text.LastIndexOf(' ', currentIndex + targetLineLength);

                // If no space is found within the target range, extend the search slightly beyond it
                if (wrapAt == -1 || wrapAt < currentIndex)
                {
                    wrapAt = text.IndexOf(' ', currentIndex + targetLineLength);
                }

                // If no space is found at all, wrap at the max line length
                if (wrapAt == -1)
                {
                    wrapAt = currentIndex + targetLineLength;
                }

                // Add the wrapped line to the result
                wrappedText.AppendLine(text.Substring(currentIndex, wrapAt - currentIndex).Trim());

                // Move the current index to the next chunk
                currentIndex = wrapAt + 1; // Move past the space for the next line
            }

            return wrappedText.ToString().TrimEnd(); // Remove trailing newline
        }
    }
}
