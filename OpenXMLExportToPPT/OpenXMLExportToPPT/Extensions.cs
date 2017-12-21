using System;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using System.IO;

namespace DocumentFormat.OpenXml.Packaging.Extensions
{
    public static class PowerpointExtensions
    {
        internal static Slide InsertSlide(this PresentationPart presentationPart, string layoutName)
        {
            //1) create the slide
            Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
            
            ////2) specify non-visual properties of the new slide
            //NonVisualGroupShapeProperties nonVisualProperties = slide.CommonSlideData.ShapeTree.AppendChild(new NonVisualGroupShapeProperties());
            //nonVisualProperties.NonVisualDrawingProperties = new NonVisualDrawingProperties() { Id = 1, Name = "" };
            //nonVisualProperties.NonVisualGroupShapeDrawingProperties = new NonVisualGroupShapeDrawingProperties();
            //nonVisualProperties.ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();

            ////3) Specify the group shape properties of the new slide.
            //slide.CommonSlideData.ShapeTree.AppendChild(new GroupShapeProperties());

            //4) create slide part for the new slide inside the presentation part
            SlidePart sPart = presentationPart.AddNewPart<SlidePart>();
            slide.Save(sPart);

            //5) set the slidelayout
            SlideMasterPart smPart = presentationPart.SlideMasterParts.First();
            SlideLayoutPart slPart = smPart.SlideLayoutParts.SingleOrDefault(sl => sl.SlideLayout.CommonSlideData.Name.Value.Equals(layoutName));
            sPart.AddPart<SlideLayoutPart>(slPart);

            //sPart.CommonSlideData = (CommonSlideData)smPart.SlideLayoutParts.SingleOrDefault(
            //    sl => sl.SlideLayout.CommonSlideData.Name.Value.Equals(layoutName)).SlideLayout.CommonSlideData.Clone();

            //6) set the slideid
            UInt32 slideId = 256U;
            slideId += Convert.ToUInt32(presentationPart.Presentation.SlideIdList.Count());
            SlideId newSlideId = presentationPart.Presentation.SlideIdList.AppendChild<SlideId>(new SlideId());
            newSlideId.Id = slideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(sPart);

            return GetSlideByRelationshipId(presentationPart, newSlideId.RelationshipId);
        }

        private static Slide GetSlideByRelationshipId(PresentationPart presentationPart, StringValue relId)
        {
            SlidePart slidePart = presentationPart.GetPartById(relId) as SlidePart;
            if (slidePart != null)
            {
                return slidePart.Slide;
            }
            else
            {
                return null;
            }
        }

        internal static Picture AddPicture(this Slide slide, Shape referingShape, string imageFile)
        {
            Picture picture = new Picture();

            string embedId = string.Empty;
            UInt32Value picId = 10001U;
            string name = string.Empty;

            if (slide.Elements<Picture>().Count() > 0)
            {
                picId = ++slide.Elements<Picture>().ToList().Last().NonVisualPictureProperties.NonVisualDrawingProperties.Id;
            }
            name = "image" + picId.ToString();
            embedId = "rId" + (slide.Elements<Picture>().Count() + 915).ToString(); // some value

            NonVisualPictureProperties nonVisualPictureProperties = new NonVisualPictureProperties()
            {
                NonVisualDrawingProperties = new NonVisualDrawingProperties() { Name = name, Id = picId, Title = name },
                NonVisualPictureDrawingProperties = new NonVisualPictureDrawingProperties() { PictureLocks = new Drawing.PictureLocks() { NoChangeAspect = true } },
                ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties() { UserDrawn = true }
            };

            BlipFill blipFill = new BlipFill() { Blip = new Drawing.Blip() { Embed = embedId } };
            Drawing.Stretch stretch = new Drawing.Stretch() { FillRectangle = new Drawing.FillRectangle() };
            blipFill.Append(stretch);

            ShapeProperties shapeProperties = new ShapeProperties()
            {
                Transform2D = new Drawing.Transform2D()
                {
                    Offset = new Drawing.Offset() { X = 1554691, Y = 1600200 },
                    Extents = new Drawing.Extents() { Cx = 6034617, Cy = 4525963 }
                }
            };
            Drawing.PresetGeometry presetGeometry = new Drawing.PresetGeometry() { Preset = Drawing.ShapeTypeValues.Rectangle };
            Drawing.AdjustValueList adjustValueList = new Drawing.AdjustValueList();

            presetGeometry.Append(adjustValueList);
            shapeProperties.Append(presetGeometry);
            picture.Append(nonVisualPictureProperties);
            picture.Append(blipFill);
            picture.Append(shapeProperties);

            slide.CommonSlideData.ShapeTree.Append(picture);

            // Add Image part
            slide.AddImagePart(embedId, imageFile);

            slide.Save();
            return picture;
        }

        private static void AddImagePart(this Slide slide, string relationshipId, string imageFile)
        {
            ImagePart imgPart = slide.SlidePart.AddImagePart(GetImagePartType(imageFile), relationshipId);
            using (FileStream imgStream = File.Open(imageFile, FileMode.Open))
            {
                imgPart.FeedData(imgStream);
            }
        }

        private static ImagePartType GetImagePartType(string imageFile)
        {
            string[] imgFileSplit = imageFile.Split('.');
            string imgExtension = imgFileSplit.ElementAt(imgFileSplit.Count() - 1).ToString().ToLower();
            if (imgExtension.Equals("png"))
                imgExtension = "png";
            return (ImagePartType)Enum.Parse(typeof(ImagePartType), imgExtension, true);
        }

        public static bool CopyPresentation(string fileNameWithExtension, string copyToPath)
        {
            try
            {
                File.Copy(fileNameWithExtension, copyToPath);
                return true;
            }
            catch (Exception)
            {
                throw;
            }
        }

        }

    }
