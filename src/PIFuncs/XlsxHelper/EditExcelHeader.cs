using ClosedXML.Excel;
using System;
using System.Linq;
using System.Text;

namespace PIFunc.XlsxHelper
{
    public class EditExcelHeader
    {
        public static void ModifyHeaderSection(XLWorkbook workbook, string docId, string procedureRef, string revisionNo, string revisionDate, string fileName)
        {
            try
            {
                foreach (IXLWorksheet worksheet in workbook.Worksheets)
                {
                    foreach (IXLRow row in worksheet.Rows(1, 7))
                    {
                        int docIdColumn = -1;
                        int procedureRefColumn = -1;
                        int revisionNoColumn = -1;
                        int revisionDateColumn = -1;
                        int documentNameColumn = -1;
                        int copyNoColumn = -1;
                        int controlledStampColumn = -1;
                        int pageColumn = -1;

                        foreach (IXLCell cell in row.Cells())
                        {
                            string cellValue = cell.GetString();

                            if (cellValue.Contains("DOC ID"))
                            {
                                UpdateAdjacentCell(worksheet, cell, docId, false);
                                docIdColumn = cell.Address.ColumnNumber;
                            }
                            else if (cellValue.Contains("PROCEDURE REF"))
                            {
                                UpdateAdjacentCell(worksheet, cell, procedureRef, false);
                                procedureRefColumn = cell.Address.ColumnNumber;
                            }
                            else if (cellValue.Contains("REVISION NO"))
                            {
                                UpdateAdjacentCell(worksheet, cell, revisionNo, false);
                                revisionNoColumn = cell.Address.ColumnNumber;
                            }
                            else if (cellValue.Contains("REVISION DATE"))
                            {
                                UpdateAdjacentCell(worksheet, cell, revisionDate, false);
                                revisionDateColumn = cell.Address.ColumnNumber;
                            }
                            else if (cellValue.Contains("Document Name"))
                            {
                                UpdateAdjacentCell(worksheet, cell, fileName, true); // Bold the Document Name
                                documentNameColumn = cell.Address.ColumnNumber;
                            }
                            else if (cellValue.Contains("COPY NO."))
                            {
                                copyNoColumn = cell.Address.ColumnNumber;
                            }
                            else if (cellValue.Contains("Controlled, If stamped in red."))
                            {
                                controlledStampColumn = cell.Address.ColumnNumber;
                            }
                            else if (cellValue.Contains("PAGE:"))
                            {
                                pageColumn = cell.Address.ColumnNumber;
                            }
                        }

                        // Handle "DOC ID" to "PROCEDURE REF" merge logic
                        if (docIdColumn != -1 && procedureRefColumn != -1 && docIdColumn < procedureRefColumn - 1)
                        {
                            MergeBetweenCells(worksheet, row.RowNumber(), docIdColumn + 1, procedureRefColumn - 1);
                        }

                        // Handle "PROCEDURE REF" to the last used column merge logic
                        if (procedureRefColumn != -1)
                        {
                            int lastColumn = GetLastUsedColumnInRow(row);

                            // Ensure the merge starts after the "PROCEDURE REF" column and excludes it
                            if (procedureRefColumn < lastColumn)
                            {
                                MergeBetweenCells(worksheet, row.RowNumber(), procedureRefColumn + 1, lastColumn);
                            }
                        }

                        // Handle "REVISION NO" to "REVISION DATE" merge logic
                        if (revisionNoColumn != -1 && revisionDateColumn != -1 && revisionNoColumn < revisionDateColumn - 1)
                        {
                            MergeBetweenCells(worksheet, row.RowNumber(), revisionNoColumn + 1, revisionDateColumn - 1);
                        }

                        // Handle "REVISION DATE" to the last used column merge logic
                        if (revisionDateColumn != -1)
                        {
                            int lastColumn = GetLastUsedColumnInRow(row);

                            // Ensure the merge starts after the "REVISION DATE" column and excludes it
                            if (revisionDateColumn < lastColumn)
                            {
                                MergeBetweenCells(worksheet, row.RowNumber(), revisionDateColumn + 1, lastColumn);
                            }
                        }

                        // Handle "Document Name" to the last column merge logic
                        if (documentNameColumn != -1)
                        {
                            int lastColumn = GetLastUsedColumnInRow(row);
                            if (documentNameColumn < lastColumn)
                            {
                                // Merge cells after "Document Name"
                                MergeBetweenCells(worksheet, row.RowNumber(), documentNameColumn + 1, lastColumn, true);

                                // Ensure the text is bold in the merged range
                                var mergedRange = worksheet.Range(row.RowNumber(), documentNameColumn + 1, row.RowNumber(), lastColumn);
                                mergedRange.Style.Font.Bold = true;
                            }
                        }

                        // Merge between "COPY NO." and "Controlled, If stamped in red."
                        if (copyNoColumn != -1 && controlledStampColumn != -1 && copyNoColumn < controlledStampColumn - 1)
                        {
                            MergeBetweenCells(worksheet, row.RowNumber(), copyNoColumn + 1, controlledStampColumn - 1);
                        }

                        // Merge between "Controlled, If stamped in red." and "PAGE: of"
                        if (controlledStampColumn != -1 && pageColumn != -1 && controlledStampColumn < pageColumn - 1)
                        {
                            MergeBetweenCells(worksheet, row.RowNumber(), controlledStampColumn + 1, pageColumn - 1);
                        }
                    }
                }

                workbook.Save();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to get the last used column number of a row, considering all cells, including empty ones
        private static int GetLastUsedColumnInRow(IXLRow row)
        {
            return row.Cells().Select(cell =>
            {
                var mergedRange = cell.MergedRange();
                return mergedRange != null ? mergedRange.LastCell().Address.ColumnNumber : cell.Address.ColumnNumber;
            }).Max();
        }

        // Method to merge cells in a given range, but retain the content of the first cell in that range
        // Added parameter to center text for "Document Name" row
        private static void MergeBetweenCells(IXLWorksheet worksheet, int rowNumber, int startColumn, int endColumn, bool centerText = false)
        {
            if (startColumn <= endColumn)
            {
                // Get the first cell in the range
                IXLCell firstCell = worksheet.Cell(rowNumber, startColumn);

                // Merge the cells in the specified range
                var mergedRange = worksheet.Range(rowNumber, startColumn, rowNumber, endColumn).Merge();

                // Retain the content of the first cell in the merged range
                worksheet.Cell(rowNumber, startColumn).Value = firstCell.GetString();

                // Apply borders to the merged cells
                mergedRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                mergedRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                // Center the text if needed (e.g., for "Document Name")
                if (centerText)
                {
                    mergedRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    mergedRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                }
            }
        }

        // Helper method to update the adjacent cell with the new value, wrap text, and apply formatting
        private static void UpdateAdjacentCell(IXLWorksheet worksheet, IXLCell cell, string newValue, bool isBold)
        {
            int columnNumber = cell.Address.ColumnNumber;
            int rowNumber = cell.Address.RowNumber;

            IXLCell nextCell = worksheet.Cell(rowNumber, columnNumber + 1);

            string wrappedText = WrapText(newValue, 50); // Assume a max of 50 characters per line
            nextCell.Value = wrappedText;

            nextCell.Style.Font.FontName = "Times New Roman";
            nextCell.Style.Font.FontSize = 12;

            if (isBold)
            {
                // Capitalize the text for bold formatting
                nextCell.Value = wrappedText.ToUpper();
                nextCell.Style.Font.Bold = true;
                nextCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                nextCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            }
            else
            {
                // Just set text without bold formatting
                nextCell.Style.Font.Bold = false;
                nextCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                nextCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            }

            nextCell.Style.Alignment.WrapText = true;
        }

        // Function to wrap text based on target line length
        private static string WrapText(string text, int targetLineLength)
        {
            StringBuilder wrappedText = new StringBuilder();
            int currentIndex = 0;

            while (currentIndex < text.Length)
            {
                int remainingLength = text.Length - currentIndex;
                if (remainingLength <= targetLineLength)
                {
                    wrappedText.AppendLine(text.Substring(currentIndex).Trim());
                    break;
                }

                int wrapAt = text.LastIndexOf(' ', currentIndex + targetLineLength);

                if (wrapAt == -1 || wrapAt < currentIndex)
                {
                    wrapAt = text.IndexOf(' ', currentIndex + targetLineLength);
                }

                if (wrapAt == -1)
                {
                    wrapAt = currentIndex + targetLineLength;
                }

                wrappedText.AppendLine(text.Substring(currentIndex, wrapAt - currentIndex).Trim());
                currentIndex = wrapAt + 1;
            }

            return wrappedText.ToString().TrimEnd();
        }

        // Helper method to convert column number to column letter
        private static string ColumnNumberToLetter(int columnNumber)
        {
            string columnName = "";
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = (char)(65 + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }
            return columnName;
        }
    }
}
