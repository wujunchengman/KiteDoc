using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using KiteDoc.Enum;
using KiteDoc.Utils;
using System;
using System.IO;
using System.Linq;
using SixLabors.ImageSharp;


using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml.Wordprocessing;

namespace KiteDoc
{
    /// <summary>
    /// 替换文本为图片
    /// </summary>
    public static class DocReplacePictureExtension
    {
        public static int Replace(this WordprocessingDocument doc,string oldString, string fileName, ImageType imageType, int width = 18, int height = -1)
        {
            var count = 0;

            var elements = doc.FindAllTextElement();
            var waitReplace = DocElementUtils.FindRun(elements, oldString);

            if (waitReplace.Any()) {
                if (!File.Exists(fileName))
                {
                    throw new Exception($"未在该路径({fileName})找到图片");
                }
                MainDocumentPart mainPart = doc.MainDocumentPart;
                ImagePartType imagePartType = ImagePartType.Png;
                switch (imageType)
                {
                    case ImageType.Png:
                        imagePartType = ImagePartType.Png;
                        break;
                    case ImageType.Jpeg:
                        imagePartType = ImagePartType.Jpeg;
                        break;
                    case ImageType.Bmp:
                        imagePartType = ImagePartType.Bmp;
                        break;
                    case ImageType.Gif:
                        imagePartType = ImagePartType.Gif;
                        break;
                    case ImageType.Icon:
                        imagePartType = ImagePartType.Icon;
                        break;
                }

                ImagePart imagePart = mainPart.AddImagePart(imagePartType);


                double rate = default;
                // 将图片写入Word
                using (FileStream stream = new FileStream(fileName, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                }
                // 再次读取计算宽高比
                // 这里因为流写过一次就变成空了，需要深入了解后再优化
                using (FileStream stream = new FileStream(fileName, FileMode.Open))
                {
                    var imageInfo = Image.Identify(stream);
                    rate = (double)imageInfo.Width / (double)imageInfo.Height;
                }


                string relationshipId = mainPart.GetIdOfPart(imagePart);

                long cx = 360000L * width; //360000L = 1厘米
                long cy = default;
                if (height == -1)
                {
                    cy = (long)(360000L * width / rate);
                }
                else
                {
                    cy = 360000L * height;
                }

                // Define the reference of the image.
                var element =
                     new Drawing(
                         new DW.Inline(
                             new DW.Extent() { Cx = cx, Cy = cy },
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
                                 Name = Path.GetFileNameWithoutExtension(fileName)
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
                                                 new A.Extents() { Cx = cx, Cy = cy }),
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

                var run = new Run(element);


                foreach (var item in waitReplace)
                {
                    // 获取图片，插入到对应的位置
                    item.First().InsertAfterSelf( (Run)run.Clone() );

                    foreach (var waitDelete in item)
                    {
                        waitDelete.Remove();
                    }

                    count++;
                }

            }



            return count;
        }
    }
}
