using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
namespace OpenXMLExportToPPT
{
    public partial class Report1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }
        protected void btn_submit_Click(object sender, EventArgs e)
        {
            string pptFilePath = @"C:\pics";
            string themeFilePath = @"C:\ZainPPTtemplate.pptx";
            ExportToPPT(pptFilePath, themeFilePath);


        }
        public void ExportToPPT(string folderPath, string themeFilePath)
        {
            string pptFilePath = folderPath + @"\Report.pptx";
            if (CopyPresentation(themeFilePath, pptFilePath))
            {
                using (PresentationDocument presentationDocument = PresentationDocument.Open(pptFilePath, true))
                {
                    //string xml = (new StreamReader(presentationDocument.PresentationPart.GetStream())).ReadToEnd();
                    DirectoryInfo imagesDirectory = new DirectoryInfo(folderPath);
                    var presentationPart = presentationDocument.PresentationPart;
                    //MMS:2
                    var templatePart = GetSlidePartsInOrder(presentationPart).ElementAt(4);

                    foreach (FileInfo file in imagesDirectory.GetFiles("*.png"))
                    {
                        var newSlidePart = CloneSlide(templatePart);
                        AppendSlide(presentationPart,newSlidePart);
                        Slide slide = newSlidePart.Slide;

                        //MMS:1
                        Picture pic = newSlidePart.Slide.Descendants<Picture>().First();
                        AddImagePart(slide, pic.BlipFill.Blip.Embed.Value, file.FullName);

                        //using (FileStream imgStream = File.Open(file.FullName, FileMode.Open))
                        //{
                        //    ip.FeedData(imgStream);
                        //}
                         //string xml = (new StreamReader(newSlidePart.GetStream())).ReadToEnd();
                        //Shape shape = slide.CommonSlideData.ShapeTree.Elements<Shape>().FirstOrDefault(
                        //sh => sh.NonVisualShapeProperties.NonVisualDrawingProperties.Name.Value.ToLower().Equals("Picture Placeholder 5".ToLower()));
                        //Picture pic = AddPicture(slide, shape, file.FullName);
                        //slide.CommonSlideData.ShapeTree.RemoveChild<Shape>(shape);
                        slide.Save();
                         //string xml2 = (new StreamReader(newSlidePart.GetStream())).ReadToEnd();
                        
                    }
                    //presentationDocument.PresentationPart.DeletePart(templatePart);
                    string xml = (new StreamReader(presentationDocument.PresentationPart.GetStream())).ReadToEnd();
                    presentationDocument.PresentationPart.Presentation.Save();
                    presentationDocument.Close();
                }
            }
        }
        private static bool CopyPresentation(string fileNameWithExtension, string copyToPath)
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
        private static IEnumerable<SlidePart> GetSlidePartsInOrder(PresentationPart presentationPart)
        {
            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

            return slideIdList.ChildElements
                .Cast<SlideId>()
                .Select(x => presentationPart.GetPartById(x.RelationshipId))
                .Cast<SlidePart>();
        }
        private static SlidePart CloneSlide(SlidePart templatePart)
        {
            // find the presentationPart: makes the API more fluent
            var presentationPart = templatePart.GetParentParts()
                .OfType<PresentationPart>()
                .Single();

            // clone slide contents
            Slide currentSlide = (Slide)templatePart.Slide.CloneNode(true);
            var slidePartClone = presentationPart.AddNewPart<SlidePart>();
            currentSlide.Save(slidePartClone);

            // copy layout part
            slidePartClone.AddPart(templatePart.SlideLayoutPart);

            return slidePartClone;
        }
        private static string AppendSlide(PresentationPart presentationPart, SlidePart newSlidePart)
        {
            //MMS:3

            //get slides id list
            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

            // find the highest id
            uint maxSlideId = slideIdList.ChildElements
                .Cast<SlideId>()
                .Max(x => x.Id.Value);

            //create new slide id based on max id
            uint newId = maxSlideId + 1;
            
            //add new slide id item at the second place in the list
            SlideId newSlideId = new SlideId();
            slideIdList.InsertAt(newSlideId, 1);
            newSlideId.Id = newId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(newSlidePart);
            //newSlideId.RelationshipId.Value = "rId2";
            
            
            return newSlideId.RelationshipId;
        }
        private static Picture AddPicture( Slide slide, Shape referingShape, string imageFile)
        {
            Picture picture = new Picture();

            string embedId = string.Empty;
            DocumentFormat.OpenXml.UInt32Value picId = 10001U;
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
                NonVisualPictureDrawingProperties = new NonVisualPictureDrawingProperties() { PictureLocks = new DocumentFormat.OpenXml.Drawing.PictureLocks() { NoChangeAspect = true } },
                ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties() { UserDrawn = true }
            };

            BlipFill blipFill = new BlipFill() { Blip = new DocumentFormat.OpenXml.Drawing.Blip() { Embed = embedId } };
            DocumentFormat.OpenXml.Drawing.Stretch stretch = new DocumentFormat.OpenXml.Drawing.Stretch() { FillRectangle = new DocumentFormat.OpenXml.Drawing.FillRectangle() };
            blipFill.Append(stretch);

            ShapeProperties shapeProperties = new ShapeProperties()
            {
                Transform2D = new DocumentFormat.OpenXml.Drawing.Transform2D()
                {
                    //MMS:4
                    Offset = new DocumentFormat.OpenXml.Drawing.Offset() { X = 1565275, Y = 612775 },// { X = 457200L, Y = 1124000L }
                    Extents = new DocumentFormat.OpenXml.Drawing.Extents() { Cx = 5486400, Cy = 4114800 }//{ Cx = 8229600L, Cy = 5029200L }
                }

                            //<a:off x="1565275" y="612775" />
                            //<a:ext cx="5486400" cy="4114800" />
            };
            DocumentFormat.OpenXml.Drawing.PresetGeometry presetGeometry = new DocumentFormat.OpenXml.Drawing.PresetGeometry() { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle };
            DocumentFormat.OpenXml.Drawing.AdjustValueList adjustValueList = new DocumentFormat.OpenXml.Drawing.AdjustValueList();

            presetGeometry.Append(adjustValueList);
            shapeProperties.Append(presetGeometry);
            picture.Append(nonVisualPictureProperties);
            picture.Append(blipFill);
            picture.Append(shapeProperties);

            slide.CommonSlideData.ShapeTree.Append(picture);

            // Add Image part
            AddImagePart(slide,embedId, imageFile);

            slide.Save();
            return picture;
        }
        private static void AddImagePart( Slide slide, string relationshipId, string imageFile)
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
    }
}