using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace WordProcessingEx
{
    class Program
    {
        static void Main(string[] args)
        {
            // Apply the Heading 3 style to a paragraph.   
            string fileName = @"C:\Users\Public\Documents\demo.docx";
            string imgName = @"C:\Users\Public\Documents\testqr.jpg";
            using (WordprocessingDocument myDocument = WordprocessingDocument.Open(fileName, true))

            {
                MainDocumentPart mainPart = myDocument.MainDocumentPart;
                HeaderPart headerPart = mainPart.HeaderParts.FirstOrDefault();
                if (headerPart == null)
                { headerPart = mainPart.AddNewPart<HeaderPart>(); }

                if (headerPart.Header == null)
                {
                    string newHeaderText = "New header via Open XML Format SDK 2.0 classes";

                    PositionalTab pTab = new PositionalTab()
                    {
                        Alignment = AbsolutePositionTabAlignmentValues.Center,
                        RelativeTo = AbsolutePositionTabPositioningBaseValues.Margin,
                        Leader = AbsolutePositionTabLeaderCharValues.None
                    };
                    headerPart.Header = new Header(
                    new Paragraph(
                      new ParagraphProperties(
                        new ParagraphStyleId() { Val = "Header" }),
                      new Run(pTab,
                        new Text(newHeaderText))
                    )
                  );
                    string rId = mainPart.GetIdOfPart(headerPart);
                    foreach (SectionProperties sectProperties in mainPart.Document.Descendants<SectionProperties>())
                    {
                        foreach (HeaderReference headerReference in sectProperties.Descendants<HeaderReference>())
                            sectProperties.RemoveChild(headerReference);
                        HeaderReference newHeaderReference = new HeaderReference() { Id = rId, Type = HeaderFooterValues.Default };
                        sectProperties.Append(newHeaderReference);
                    }
                    HeaderReference headerReference1 = new HeaderReference() { Type = HeaderFooterValues.Default, Id = rId };
                    IEnumerable<SectionProperties> sections = mainPart.Document.Body.Elements<SectionProperties>();

                    foreach (var section in sections)
                    {
                        // Delete existing references to headers and footers
                        section.RemoveAllChildren<HeaderReference>();
                        // Create the new header and footer reference node
                        section.PrependChild<HeaderReference>(new HeaderReference() { Id = rId });
                        
                    }
                }
                Header hd = headerPart.Header;
                ImagePart img = headerPart.AddImagePart(ImagePartType.Jpeg);
                using (FileStream stream = new FileStream(imgName, FileMode.Open))
                {
                    img.FeedData(stream);
                }
                AddImageToHeader(hd, headerPart.GetIdOfPart(img));
                ImagePart img2 = mainPart.AddImagePart(ImagePartType.Jpeg);
                using (FileStream stream = new FileStream(imgName, FileMode.Open))
                {
                    img2.FeedData(stream);
                }
                AddImageToBody(myDocument, mainPart.GetIdOfPart(img2));

            }
            Console.WriteLine("All done. Press a key.");
            //Console.ReadKey();
        }

        private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId)
        {
            // Define the reference of the image.
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = 990000L, Cy = 792000L },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            // Append the reference to body, the element should be in a Run.
            wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
        }
        private static void AddImageToHeader(Header hd, string relationshipId)
        {
            // Define the reference of the image.
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = 990000L, Cy = 792000L },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            // Append the reference to body, the element should be in a Run.
            hd.AppendChild(new Paragraph(new Run(element)));
        }
    }
}