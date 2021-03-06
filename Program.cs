using CommandLine;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using QRCoder;
using System.Drawing;
using System.Drawing.Imaging;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace QRWordInjector
{
    class Program
    {
        public class Options
        {
            [Option('f', "file", Required = true, HelpText = "Set file to add QR code.")]
            public string filename { get; set; }

            [Option('t', "text", Required = true, HelpText = "Set text for QR code.")]
            public string QRCodeData { get; set; }
        }
        static void Main(string[] args)
        {
            string stringtoencode = @"placeholder";
            string fileName = @"placeholder";
            Parser.Default.ParseArguments<Options>(args)
               .WithParsed<Options>(o =>
               {
                   fileName = o.filename;
                   stringtoencode = o.QRCodeData;
               })
               .WithNotParsed<Options>(o =>
               { Environment.Exit(1); });
            using (WordprocessingDocument myDocument = WordprocessingDocument.Open(fileName, true))

            {

                MainDocumentPart mainPart = myDocument.MainDocumentPart;
                HeaderPart headerPart = mainPart.HeaderParts.FirstOrDefault();
                headerPart = mainPart.AddNewPart<HeaderPart>();
                string rId = mainPart.GetIdOfPart(headerPart);
                Header hd = new Header();
                headerPart.Header = hd;
                HeaderReference headerReference1 = new HeaderReference() { Type = HeaderFooterValues.First, Id = rId };
                /*foreach (SectionProperties sectProperties in mainPart.Document.Descendants<SectionProperties>().First())
                {
                    DocGrid docGrid1 = sectProperties.GetFirstChild<DocGrid>();
                    TitlePage titlePage1 = new TitlePage();
                    sectProperties.InsertBefore(titlePage1, docGrid1);
                    sectProperties.Append(headerReference1);
                }*/
                SectionProperties sectProperties = mainPart.Document.Descendants<SectionProperties>().First();
                DocGrid docGrid1 = sectProperties.GetFirstChild<DocGrid>();
                TitlePage titlePage1 = new TitlePage();
                sectProperties.InsertBefore(titlePage1, docGrid1);
                sectProperties.Append(headerReference1);

                ImagePart img = headerPart.AddImagePart(ImagePartType.Jpeg);
                QRCodeGenerator qrGenerator = new QRCodeGenerator();
                QRCodeData qrCodeData = qrGenerator.CreateQrCode(stringtoencode, QRCodeGenerator.ECCLevel.Q);
                QRCode qrCode = new QRCode(qrCodeData);
                Bitmap qrCodeImage = qrCode.GetGraphic(20);
                using (MemoryStream stream1 = new MemoryStream())
                {
                    qrCodeImage.Save(stream1, ImageFormat.Jpeg);
                    stream1.Seek(0, SeekOrigin.Begin);
                    img.FeedData(stream1);
                }
                AddImageToHeader(hd, headerPart.GetIdOfPart(img));
                foreach (HeaderPart hdp in mainPart.HeaderParts) { if (hdp != headerPart) {
                        foreach (Paragraph p in hdp.Header.Descendants<Paragraph>())
                            headerPart.Header.AppendChild(p.CloneNode(true));
                    } }
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
            //hd.AppendChild(new Paragraph(new Run(element)));
            Paragraph p = hd.GetFirstChild<Paragraph>();
            hd.InsertBefore(new Paragraph(new Run(element)), p);
        }
    }
}