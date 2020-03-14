using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System;
using System.Drawing;

namespace WordManulipation
{
    public class ImageData
    {
        public string FileName = string.Empty;
        public byte[] BinaryData;
        public Stream DataStream => new MemoryStream(BinaryData);
        public ImagePartType ImageType
        {
            get
            {
                var ext = Path.GetExtension(FileName).TrimStart('.').ToLower();
                switch (ext)
                {
                    case "jpg":
                        return ImagePartType.Jpeg;
                    case "png":
                        return ImagePartType.Png;
                    case "":
                        return ImagePartType.Gif;
                    case "bmp":
                        return ImagePartType.Bmp;
                }
                throw new ApplicationException($"Unsupported image type: {ext}");
            }
        }
        public int SourceWidth;
        public int SourceHeight;
        public decimal Width;
        public decimal Height;
        public long WidthInEMU => Convert.ToInt64(Width * CM_TO_EMU);
        public long HeightInEMU => Convert.ToInt64(Height * CM_TO_EMU);
        private const decimal INCH_TO_CM = 2.54M;
        private const decimal CM_TO_EMU = 360000M;
        public string ImageName;
        public ImageData(string fileName, byte[] data, int dpi = 300)
        {
            FileName = fileName;
            BinaryData = data;
            Bitmap img = new Bitmap(new MemoryStream(data));
            SourceWidth = img.Width;
            SourceHeight = img.Height;
            Width = ((decimal)SourceWidth) / dpi * INCH_TO_CM;
            Height = ((decimal)SourceHeight) / dpi * INCH_TO_CM;
            ImageName = $"IMG_{Guid.NewGuid().ToString().Substring(0, 8)}";
        }
        public ImageData(string fileName, int dpi = 300) :
            this(fileName, File.ReadAllBytes(fileName), dpi)
        {
        }
    }
    public class DocxImgHelper
    {
        public static Run GenerateImageRun(WordprocessingDocument wordDoc, ImageData img)
        {
            MainDocumentPart mainPart = wordDoc.MainDocumentPart;

            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
            var relationshipId = mainPart.GetIdOfPart(imagePart);
            imagePart.FeedData(img.DataStream);

            // Define the reference of the image.
            var element =
                 new Drawing(
                     new DW.Inline(
                         //Size of image, unit = EMU(English Metric Unit)
                         //1 cm = 360000 EMUs
                         new DW.Extent() { Cx = img.WidthInEMU, Cy = img.HeightInEMU },
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
                             Name = img.ImageName
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
                                             Name = img.FileName
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
                                             new A.Extents()
                                             {
                                                 Cx = img.WidthInEMU,
                                                 Cy = img.HeightInEMU
                                             }),
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
            return new Run(element);
        }
    }
}
