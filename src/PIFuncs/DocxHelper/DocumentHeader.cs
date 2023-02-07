using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Collections.Generic;

namespace PIFunc.DocxHelper
{
    internal class DocumentHeader
    {
        public static void AddMetadata(WordprocessingDocument document, Table table)
        {

            MainDocumentPart mainDocumentPart = document.MainDocumentPart;
            // Delete the existing header and footer parts
            mainDocumentPart.DeleteParts(mainDocumentPart.HeaderParts);
            HeaderPart headerPart = mainDocumentPart.AddNewPart<HeaderPart>();

            string headerPartId = mainDocumentPart.GetIdOfPart(headerPart);

            Header header1 = new Header() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 wp14" } };
            header1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            header1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            header1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            header1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            header1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            header1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            header1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            header1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            header1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            header1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            header1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            header1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            header1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            header1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            header1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00164C17", RsidRunAdditionDefault = "00164C17" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

            paragraphProperties1.Append(paragraphStyleId1);

            Run run1 = new Run();
            Text text1 = new Text();
            text1.Text = "This my sample header from openxml docs";

           // run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            header1.Append(paragraph1);
            header1.Append(table);
            header1.Append(CreatePageNumbering());
            headerPart.Header = header1;


            // Get SectionProperties and Replace HeaderReference and FooterRefernce with new Id
            IEnumerable<SectionProperties> sections = mainDocumentPart.Document.Body.Elements<SectionProperties>();

            foreach (var section in sections)
            {
                // Delete existing references to headers and footers
                section.RemoveAllChildren<HeaderReference>();

                // Create the new header and footer reference node
                section.PrependChild<HeaderReference>(new HeaderReference() { Id = headerPartId });
            }
        }

        public static SdtBlock CreatePageNumbering()
        {
            SdtBlock objSdtBlock_1 = new SdtBlock();
            SdtContentBlock objSdtContentBlock_1 =
                new SdtContentBlock();
            SdtBlock objSdtBlock_2 = new SdtBlock();
            SdtContentBlock objSdtContentBlock_2 =
                new SdtContentBlock();
            Paragraph objParagraph_1 = new Paragraph();
            ParagraphProperties objParagraphProperties =
                new ParagraphProperties();
            ParagraphStyleId objParagraphStyleId =
                new ParagraphStyleId() { Val = "Header" };
            objParagraphProperties.Append(objParagraphStyleId);
            Justification objJustification =
                new Justification() { Val = JustificationValues.Right };
            objParagraphProperties.Append(objJustification);
            objParagraph_1.Append(objParagraphProperties);
            Run objRun_1 = new Run();
            Text objText_1 = new Text();
            objText_1.Text = "Page ";
            objRun_1.Append(objText_1);
            objParagraph_1.Append(objRun_1);
            Run objRun_2 = new Run();
            FieldChar objFieldChar_1 =
                new FieldChar()
                { FieldCharType = FieldCharValues.Begin };
            objRun_2.Append(objFieldChar_1);
            objParagraph_1.Append(objRun_2);
            Run objRun_3 = new Run();
            FieldCode objFieldCode_1 =
                new FieldCode()
                { Space = SpaceProcessingModeValues.Preserve };
            objFieldCode_1.Text = "PAGE ";
            objRun_3.Append(objFieldCode_1);
            objParagraph_1.Append(objRun_3);
            Run objRun_4 = new Run();
            FieldChar objFieldChar_2 =
                new FieldChar()
                { FieldCharType = FieldCharValues.Separate };
            objRun_4.Append(objFieldChar_2);
            objParagraph_1.Append(objRun_4);
            Run objRun_5 = new Run();
            Text objText_2 = new Text();
            objText_2.Text = "2";
            objRun_5.Append(objText_2);
            objParagraph_1.Append(objRun_5);
            Run objRun_6 = new Run();
            FieldChar objFieldChar_3 =
                new FieldChar() { FieldCharType = FieldCharValues.End };
            objRun_6.Append(objFieldChar_3);
            objParagraph_1.Append(objRun_6);
            Run objRun_7 = new Run();
            Text objText_3 = new Text();
            objText_3.Text = " of  ";
            objRun_7.Append(objText_3);
            objParagraph_1.Append(objRun_7);
            Run objRun_8 = new Run();
            FieldChar objFieldChar_4 =
                new FieldChar()
                { FieldCharType = FieldCharValues.Begin };
            objRun_8.Append(objFieldChar_4);
            objParagraph_1.Append(objRun_8);
            Run objRun_9 = new Run();
            FieldCode objFieldCode_2 =
                new FieldCode()
                { Space = SpaceProcessingModeValues.Preserve };
            objFieldCode_2.Text = "NUMPAGES  ";
            objRun_9.Append(objFieldCode_2);
            objParagraph_1.Append(objRun_9);
            Run objRun_10 = new Run();
            FieldChar objFieldChar_5 =
                new FieldChar()
                { FieldCharType = FieldCharValues.Separate };
            objRun_10.Append(objFieldChar_5);
            objParagraph_1.Append(objRun_10);
            Run objRun_11 = new Run();
            Text objText_4 = new Text();
            objText_4.Text = "2";
            objRun_11.Append(objText_4);
            objParagraph_1.Append(objRun_11);
            Run objRun_12 = new Run();
            FieldChar objFieldChar_6 =
                new FieldChar() { FieldCharType = FieldCharValues.End };
            objRun_12.Append(objFieldChar_6);
            objParagraph_1.Append(objRun_12);
            objSdtContentBlock_2.Append(objParagraph_1);
            objSdtBlock_2.Append(objSdtContentBlock_2);
            objSdtContentBlock_1.Append(objSdtBlock_2);
            objSdtBlock_1.Append(objSdtContentBlock_1);

            return objSdtBlock_1;
        }
    }
}
