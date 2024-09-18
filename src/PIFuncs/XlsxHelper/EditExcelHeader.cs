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
                    // Calculate the max used range of the body (starting from row 8)
                    int maxBodyColumn = GetMaxUsedColumnInBody(worksheet);
                    Console.WriteLine($"Max Body Column: {maxBodyColumn}");

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
                        int piIndustriesLtdColumn = -1;

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
                            else if (cellValue.Contains("PI INDUSTRIES LTD") || cellValue.Contains("PI INDUSTURIES LTD"))
                            {
                                piIndustriesLtdColumn = cell.Address.ColumnNumber;
                            }
                        }

                        // Handle "DOC ID" to "PROCEDURE REF" merge logic
                        if (docIdColumn != -1 && procedureRefColumn != -1 && docIdColumn < procedureRefColumn - 1)
                        {
                            MergeBetweenCells(worksheet, row.RowNumber(), docIdColumn + 1, procedureRefColumn - 1);
                        }

                        // Handle "PROCEDURE REF" to maxBodyColumn merge logic
                        if (procedureRefColumn != -1 && procedureRefColumn + 1 <= maxBodyColumn)
                        {
                            MergeBetweenCells(worksheet, row.RowNumber(), procedureRefColumn + 1, maxBodyColumn);
                        }

                        // Handle "REVISION NO" to "REVISION DATE" merge logic
                        if (revisionNoColumn != -1 && revisionDateColumn != -1 && revisionNoColumn < revisionDateColumn - 1)
                        {
                            MergeBetweenCells(worksheet, row.RowNumber(), revisionNoColumn + 1, revisionDateColumn - 1);
                        }

                        // Handle "REVISION DATE" to maxBodyColumn merge logic
                        if (revisionDateColumn != -1 && revisionDateColumn + 1 <= maxBodyColumn)
                        {
                            MergeBetweenCells(worksheet, row.RowNumber(), revisionDateColumn + 1, maxBodyColumn);
                        }

                        // Handle "Document Name" to maxBodyColumn merge logic
                        if (documentNameColumn != -1 && documentNameColumn + 1 <= maxBodyColumn)
                        {
                            MergeBetweenCells(worksheet, row.RowNumber(), documentNameColumn + 1, maxBodyColumn, true);
                        }

                        // Handle "PAGE:" directly to maxBodyColumn merge logic
                        if (pageColumn != -1)
                        {
                            MergeBetweenCells(worksheet, row.RowNumber(), pageColumn, maxBodyColumn);
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

                        // Insert two blank rows immediately after the row containing "COPY NO.", "Controlled, If stamped in red.", and "PAGE: of"
                        if (copyNoColumn != -1 && controlledStampColumn != -1 && pageColumn != -1)
                        {
                            // Insert two rows below the current row
                            worksheet.Row(row.RowNumber() + 1).InsertRowsBelow(2);

                            // Merge all cells in the inserted rows
                            IXLRange mergedRange1 = worksheet.Range(row.RowNumber() + 1, copyNoColumn, row.RowNumber() + 1, maxBodyColumn).Merge();
                            IXLRange mergedRange2 = worksheet.Range(row.RowNumber() + 2, copyNoColumn, row.RowNumber() + 2, maxBodyColumn).Merge();

                            // Apply no borders below and on the sides of the inserted rows
                            mergedRange1.Style.Border.OutsideBorder = XLBorderStyleValues.None;
                            mergedRange1.Style.Border.InsideBorder = XLBorderStyleValues.None;
                            mergedRange2.Style.Border.OutsideBorder = XLBorderStyleValues.None;
                            mergedRange2.Style.Border.InsideBorder = XLBorderStyleValues.None;
                        }

                        // Handle "PI INDUSTRIES LTD" to maxBodyColumn merge logic
                        if (piIndustriesLtdColumn != -1)
                        {
                            // Merge "PI INDUSTRIES LTD" to maxBodyColumn and remove the right border
                            MergeBetweenCells(worksheet, row.RowNumber(), piIndustriesLtdColumn, maxBodyColumn, removeRightBorder: true);
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



        // Updated MergeBetweenCells method to include borders
        private static void MergeBetweenCells(IXLWorksheet worksheet, int rowNumber, int startColumn, int endColumn, bool centerText = false, bool removeRightBorder = false)
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

                // Remove the right border if specified
                if (removeRightBorder)
                {
                    mergedRange.Style.Border.RightBorder = XLBorderStyleValues.None;
                }

                // Center the text if needed (e.g., for "Document Name")
                if (centerText)
                {
                    mergedRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    mergedRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                }
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
        private static void MergeBetweenCellss(IXLWorksheet worksheet, int rowNumber, int startColumn, int endColumn, bool centerText = false)
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


        // Get maximum used column in the body (starting from row 8)
        private static int GetMaxUsedColumnInBody(IXLWorksheet worksheet)
        {
            int maxColumn = 0;

            // Loop through the rows in the body (starting from row 8)
            foreach (IXLRow row in worksheet.Rows(8, worksheet.LastRowUsed().RowNumber()))
            {
                // Get the last used cell in this row
                int lastUsedColumn = row.LastCellUsed()?.Address.ColumnNumber ?? 0;

                // Check for any merged ranges that might extend beyond the last used cell
                foreach (var mergedRange in worksheet.MergedRanges)
                {
                    // If the merged range intersects with the current row
                    if (mergedRange.FirstRow().RowNumber() == row.RowNumber())
                    {
                        // Find the last column in the merged range
                        int lastMergedColumn = mergedRange.LastColumn().ColumnNumber();

                        // Update the last used column if the merged range extends further
                        if (lastMergedColumn > lastUsedColumn)
                        {
                            lastUsedColumn = lastMergedColumn;
                        }
                    }
                }

                // Update max column if necessary
                if (lastUsedColumn > maxColumn)
                {
                    maxColumn = lastUsedColumn;
                }
            }

            return maxColumn;
        }


    }
}
