using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.CommandLine;
using QRCoder;
using System.Drawing;
using System.Drawing.Imaging;
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
            string imgName = @"C:\Users\Public\Documents\testqr2.jpg";
            string QRCodeData = @"The text which should be encoded.";
            using (WordprocessingDocument myDocument = WordprocessingDocument.Open(fileName, true))

            {
                MainDocumentPart mainPart = myDocument.MainDocumentPart;
                HeaderPart headerPart = mainPart.HeaderParts.FirstOrDefault();
                if (headerPart == null)
                { headerPart = mainPart.AddNewPart<HeaderPart>(); }

                if (headerPart.Header == null)
                {
                    headerPart.Header = new Header();
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
                }
                Header hd = headerPart.Header;
                ImagePart img = headerPart.AddImagePart(ImagePartType.Jpeg);
                QRCodeGenerator qrGenerator = new QRCodeGenerator();
                QRCodeData qrCodeData = qrGenerator.CreateQrCode(QRCodeData, QRCodeGenerator.ECCLevel.Q);
                QRCode qrCode = new QRCode(qrCodeData);
                Bitmap qrCodeImage = qrCode.GetGraphic(20);
                using (MemoryStream stream1 = new MemoryStream())
                {
                    qrCodeImage.Save(stream1, ImageFormat.Jpeg);
                    stream1.Seek(0, SeekOrigin.Begin);
                    img.FeedData(stream1);
                }
                AddImageToHeader(hd, headerPart.GetIdOfPart(img));
            }
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