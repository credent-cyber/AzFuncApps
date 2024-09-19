using DocumentFormat.OpenXml.Spreadsheet;
using NPOI.HPSF;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using System.Linq;
using System.Text;



namespace PIFunc.XlsxHelper
{
    public class AuditHistoryNPOI
    {
        public static void ModifyHeaderSection(XSSFWorkbook workbook, string docId, string procedureRef, string revisionNo, string revisionDate, string fileName, string filePath)
        {
            try
            {
                foreach (ISheet sheet in workbook)
                {
                    int maxBodyColumn = GetMaxUsedColumnInBody(sheet);
                    Console.WriteLine(maxBodyColumn);
                    for (int rowIndex = 0; rowIndex <= 4; rowIndex++)
                    {
                        IRow row = sheet.GetRow(rowIndex);
                        if (row == null) continue;

                        int docIdColumn = -1;
                        int procedureRefColumn = -1;
                        int revisionNoColumn = -1;
                        int revisionDateColumn = -1;
                        int documentNameColumn = -1;
                        int copyNoColumn = -1;
                        int controlledStampColumn = -1;
                        int pageColumn = -1;
                        int piIndustriesLtdColumn = -1;

                        for (int cellIndex = 0; cellIndex < row.LastCellNum; cellIndex++)
                        {
                            ICell cell = row.GetCell(cellIndex);
                            if (cell == null) continue;

                            string cellValue = cell.ToString();

                            if (cellValue.Contains("DOC ID"))
                            {
                                UpdateAdjacentCell(sheet, row, cell, docId, false);
                                docIdColumn = cell.ColumnIndex;
                            }
                            else if (cellValue.Contains("PROCEDURE REF"))
                            {
                                UpdateAdjacentCell(sheet, row, cell, procedureRef, false);
                                procedureRefColumn = cell.ColumnIndex;
                            }
                            else if (cellValue.Contains("REVISION NO"))
                            {
                                UpdateAdjacentCell(sheet, row, cell, revisionNo, false);
                                revisionNoColumn = cell.ColumnIndex;
                            }
                            else if (cellValue.Contains("REVISION DATE"))
                            {
                                UpdateAdjacentCell(sheet, row, cell, revisionDate, false);
                                revisionDateColumn = cell.ColumnIndex;
                            }
                            else if (cellValue.Trim().ToLower().Contains("document name".ToLower()))
                            {
                                UpdateAdjacentCell(sheet, row, cell, fileName, true);
                                documentNameColumn = cell.ColumnIndex;
                            }
                            else if (cellValue.Contains("COPY NO."))
                            {
                                copyNoColumn = cell.ColumnIndex;
                            }
                            else if (cellValue.Contains("Controlled, If stamped in red."))
                            {
                                controlledStampColumn = cell.ColumnIndex;
                            }
                            else if (cellValue.Contains("PAGE:"))
                            {
                                pageColumn = cell.ColumnIndex;
                            }
                            else if (cellValue.Contains("PI INDUSTRIES LTD") || cellValue.Contains("PI INDUSTURIES LTD"))
                            {
                                piIndustriesLtdColumn = cell.ColumnIndex;
                            }
                        }

                        if (docIdColumn != -1 && procedureRefColumn != -1 && docIdColumn < procedureRefColumn - 1)
                        {
                            MergeBetweenCells(sheet, row.RowNum, docIdColumn + 1, procedureRefColumn - 1);
                        }

                        if (procedureRefColumn != -1 && procedureRefColumn + 1 <= maxBodyColumn)
                        {
                            MergeBetweenCells(sheet, row.RowNum, procedureRefColumn + 1, maxBodyColumn);
                        }

                        if (revisionNoColumn != -1 && revisionDateColumn != -1 && revisionNoColumn < revisionDateColumn - 1)
                        {
                            MergeBetweenCells(sheet, row.RowNum, revisionNoColumn + 1, revisionDateColumn - 1);
                        }

                        if (revisionDateColumn != -1 && revisionDateColumn + 1 <= maxBodyColumn)
                        {
                            MergeBetweenCells(sheet, row.RowNum, revisionDateColumn + 1, maxBodyColumn);
                        }

                        if (documentNameColumn != -1 && documentNameColumn + 1 <= maxBodyColumn)
                        {
                            MergeBetweenCells(sheet, row.RowNum, documentNameColumn + 1, maxBodyColumn, true);
                        }

                        if (pageColumn != -1)
                        {
                            MergeBetweenCells(sheet, row.RowNum, pageColumn, maxBodyColumn);
                        }

                        if (copyNoColumn != -1 && controlledStampColumn != -1 && copyNoColumn < controlledStampColumn - 1)
                        {
                            MergeBetweenCells(sheet, row.RowNum, copyNoColumn + 1, controlledStampColumn - 1);
                        }

                        if (controlledStampColumn != -1 && pageColumn != -1 && controlledStampColumn < pageColumn - 1)
                        {
                            MergeBetweenCells(sheet, row.RowNum, controlledStampColumn + 1, pageColumn - 1);
                        }

                        if (copyNoColumn != -1 && controlledStampColumn != -1 && pageColumn != -1)
                        {
                            sheet.ShiftRows(row.RowNum + 1, sheet.LastRowNum, 2);
                            IRow newRow1 = sheet.CreateRow(row.RowNum + 1);
                            IRow newRow2 = sheet.CreateRow(row.RowNum + 2);
                            MergeCells(sheet, newRow1, copyNoColumn, maxBodyColumn);
                            MergeCells(sheet, newRow2, copyNoColumn, maxBodyColumn);
                        }

                        if (piIndustriesLtdColumn != -1)
                        {
                            MergeBetweenCells(sheet, row.RowNum, piIndustriesLtdColumn, maxBodyColumn, removeRightBorder: true, allBorder: true);
                        }
                    }
                }

                using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fs);
                    fs.Flush();

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        private static void MergeBetweenCells(ISheet sheet, int rowNumber, int startColumn, int endColumn, bool centerText = false, bool removeRightBorder = false, bool allBorder = false)
        {
            if (startColumn <= endColumn)
            {
                if (startColumn == endColumn)
                {
                    Console.WriteLine($"Skipping merge for single cell at column {startColumn}");
                    return;
                }

                var newRegion = new NPOI.SS.Util.CellRangeAddress(rowNumber, rowNumber, startColumn, endColumn);

                // Remove overlapping regions
                var overlappingRegions = sheet.MergedRegions
                    .Where(region => IsOverlap(region, newRegion))
                    .ToList();

                foreach (var region in overlappingRegions)
                {
                    sheet.RemoveMergedRegion(sheet.MergedRegions.IndexOf(region));
                    Console.WriteLine($"Removed overlapping region: {region.FirstColumn}-{region.LastColumn} in row {region.FirstRow}");
                }

                sheet.AddMergedRegion(newRegion);

                IRow row = sheet.GetRow(rowNumber) ?? sheet.CreateRow(rowNumber);
                ICellStyle style = sheet.Workbook.CreateCellStyle();
                style.BorderTop = BorderStyle.Thin;
                style.BorderBottom = BorderStyle.Thin;
                style.BorderLeft = BorderStyle.Thin;
                style.BorderRight = BorderStyle.Thin;

                if (removeRightBorder)
                {
                    style.BorderRight = BorderStyle.None;
                }
                if (allBorder)
                {
                    style.BorderTop = BorderStyle.Thick;
                    style.BorderLeft = BorderStyle.Thick;
                    style.BorderRight = BorderStyle.Thick;
                    style.BorderBottom = BorderStyle.Thick;
                }

                for (int col = startColumn; col <= endColumn; col++)
                {
                    ICell cell = row.GetCell(col) ?? row.CreateCell(col);
                    cell.CellStyle = style;
                    if (centerText)
                    {
                        style.Alignment = HorizontalAlignment.Center;
                        style.VerticalAlignment = VerticalAlignment.Center;
                    }
                }
            }
        }




        private static bool IsOverlap(NPOI.SS.Util.CellRangeAddress region1, NPOI.SS.Util.CellRangeAddress region2)
        {
            // Check if two regions overlap
            return region1.LastRow >= region2.FirstRow &&
                   region1.FirstRow <= region2.LastRow &&
                   region1.LastColumn >= region2.FirstColumn &&
                   region1.FirstColumn <= region2.LastColumn;
        }



        private static void UpdateAdjacentCell(ISheet sheet, IRow row, ICell cell, string newValue, bool isBold)
        {
            int columnIndex = cell.ColumnIndex;
            ICell nextCell = row.GetCell(columnIndex + 1) ?? row.CreateCell(columnIndex + 1);

            // Set the new cell value
            nextCell.SetCellValue(WrapText(newValue, 50));

            // Create a new font
            IFont font = sheet.Workbook.CreateFont();
            font.FontName = "Times New Roman";
            font.FontHeightInPoints = 12;

            font.IsBold = isBold;

            // Create a new cell style
            ICellStyle style = sheet.Workbook.CreateCellStyle();
            style.SetFont(font);
            style.WrapText = true;

            // Apply the style to the cell
            nextCell.CellStyle = style;

            //Console.WriteLine($"Applied style to cell at column {columnIndex + 1}: FontName={font.FontName}, FontHeight={font.FontHeightInPoints}, IsBold={font.IsBold}");
        }




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

        private static int GetMaxUsedColumnInBody(ISheet sheet)
        {
            int maxColumn = 0;

            for (int rowIndex = 7; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                if (row == null) continue;

                int lastUsedColumn = row.LastCellNum - 1;
                maxColumn = Math.Max(maxColumn, lastUsedColumn);
            }
            return maxColumn;
        }

        private static void MergeCells(ISheet sheet, IRow row, int startColumn, int endColumn)
        {
            // Avoid merging if start and end columns are the same
            if (startColumn == endColumn) return;

            var cellRange = new NPOI.SS.Util.CellRangeAddress(row.RowNum, row.RowNum, startColumn, endColumn);

            // Remove overlapping regions before adding the new one
            var overlappingRegions = sheet.MergedRegions
                .Where(region => IsOverlap(region, cellRange))
                .ToList();

            foreach (var region in overlappingRegions)
            {
                sheet.RemoveMergedRegion(sheet.MergedRegions.IndexOf(region));
            }

            sheet.AddMergedRegion(cellRange);
        }


        private static void CopyImages(XSSFWorkbook oldWorkbook, XSSFWorkbook newWorkbook)
        {
            // Iterate over each sheet in the old workbook
            for (int sheetIndex = 0; sheetIndex < oldWorkbook.NumberOfSheets; sheetIndex++)
            {
                var oldSheet = oldWorkbook.GetSheetAt(sheetIndex);
                var newSheet = newWorkbook.GetSheetAt(sheetIndex);

                // Get all pictures from the old workbook
                var oldPictures = oldWorkbook.GetAllPictures();

                // Create or get the drawing patriarch for the new sheet
                XSSFDrawing drawing = (XSSFDrawing)newSheet.CreateDrawingPatriarch();

                foreach (var pictureData in oldPictures)
                {
                    if (pictureData is XSSFPictureData xssfPictureData)
                    {
                        // Add picture to the new workbook
                        int pictureIndex = newWorkbook.AddPicture(xssfPictureData.Data, xssfPictureData.PictureType);

                        var anchor = new XSSFClientAnchor
                        {
                            Col1 = 0,
                            Col2 = 1,
                            Row1 = 0,
                            Row2 = 1
                        };

                        // Create picture in the new sheet
                        drawing.CreatePicture(anchor, pictureIndex);
                    }
                }
            }
        }


        private static void CopyImagesFromOldSheetToNew(XSSFWorkbook oldWorkbook, XSSFWorkbook newWorkbook)
        {
            // Copy images from old workbook to new workbook
            CopyImages(oldWorkbook, newWorkbook);
        }
    }
}
