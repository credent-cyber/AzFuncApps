// See https://aka.ms/new-console-template for more information
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;

Console.WriteLine("Hello, World!");

// Replace header in target document with header of source document.
var filepathTo = "C:\\Users\\MALAY\\Documents\\Downloads\\DM-CMRL-1.docx";

using (WordprocessingDocument
    document = WordprocessingDocument.Open(filepathTo, true))
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

    run1.Append(text1);

    paragraph1.Append(paragraphProperties1);
    paragraph1.Append(run1);

    header1.Append(paragraph1);

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

