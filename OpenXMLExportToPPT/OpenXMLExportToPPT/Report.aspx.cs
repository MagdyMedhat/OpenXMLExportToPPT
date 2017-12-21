using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Packaging.Extensions;
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
            string pptFilePath = @"C:\pics";
            string themeFilePath = @"C:\ZainPPTtemplate.pptx";
            ExportToPPT(pptFilePath, themeFilePath);
        }

        protected void btn_submit_Click(object sender, EventArgs e)
        {

        }

        //public void ExportToPPT(string folderPath, string themeFilePath)
        //{
        //    string pptFilePath = folderPath + @"\Report.pptx";
        //    if (DocumentFormat.OpenXml.Packaging.Extensions.PowerpointExtensions.CopyPresentation(themeFilePath, pptFilePath))
        //    {
        //        using (PresentationDocument presentationDocument = PresentationDocument.Open(pptFilePath, true))
        //        {
        //            DirectoryInfo imagesDirectory = new DirectoryInfo(folderPath);
        //            foreach (FileInfo file in imagesDirectory.GetFiles("*.png"))
        //            {
        //                Slide slide = presentationDocument.PresentationPart.InsertSlide("Title Only");
        //                Shape shape = slide.CommonSlideData.ShapeTree.Elements<Shape>().FirstOrDefault(
        //                    sh => sh.NonVisualShapeProperties.NonVisualDrawingProperties.Name.Value.ToLower().Equals("Content Placeholder 2".ToLower()));
        //                Picture pic = slide.AddPicture(shape, file.FullName);
        //                slide.CommonSlideData.ShapeTree.RemoveChild<Shape>(shape);
        //                slide.Save();
        //            }
        //            presentationDocument.PresentationPart.Presentation.Save();
        //            presentationDocument.Close();

        //        }
        //    }
        //}

        public void ExportToPPT(string folderPath, string themeFilePath)
        {
            string pptFilePath = folderPath + @"\Report.pptx";
            if (CopyPresentation(themeFilePath, pptFilePath))
            {
                using (PresentationDocument presentationDocument = PresentationDocument.Open(pptFilePath, true))
                {
                    DirectoryInfo imagesDirectory = new DirectoryInfo(folderPath);
                    int position = 7;
                    foreach (FileInfo file in imagesDirectory.GetFiles("*.png"))
                    {
                        Slide slide = InsertNewSlide(presentationDocument, position, file.Name);
                        Shape shape = slide.CommonSlideData.ShapeTree.Elements<Shape>().FirstOrDefault(
                            sh => sh.NonVisualShapeProperties.NonVisualDrawingProperties.Name.Value.ToLower().Equals("Content Placeholder 2".ToLower()));
                        Picture pic = slide.AddPicture(shape, file.FullName);
                        slide.CommonSlideData.ShapeTree.RemoveChild<Shape>(shape);
                        slide.Save();

                    }
                    presentationDocument.PresentationPart.Presentation.Save();
                    presentationDocument.Close();

                }
            }
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
        public static Slide InsertNewSlide(PresentationDocument presentationDocument, int position, string slideTitle)
        {
            if (presentationDocument == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            if (slideTitle == null)
            {
                throw new ArgumentNullException("slideTitle");
            }

            PresentationPart presentationPart = presentationDocument.PresentationPart;

            // Verify that the presentation is not empty.
            if (presentationPart == null)
            {
                throw new InvalidOperationException("The presentation document is empty.");
            }

            // Declare and instantiate a new slide.
            Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
            uint drawingObjectId = 1;

            // Construct the slide content.            
            // Specify the non-visual properties of the new slide.
            NonVisualGroupShapeProperties nonVisualProperties = slide.CommonSlideData.ShapeTree.AppendChild(new NonVisualGroupShapeProperties());
            nonVisualProperties.NonVisualDrawingProperties = new NonVisualDrawingProperties() { Id = 1, Name = "" };
            nonVisualProperties.NonVisualGroupShapeDrawingProperties = new NonVisualGroupShapeDrawingProperties();
            nonVisualProperties.ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();

            // Specify the group shape properties of the new slide.
            slide.CommonSlideData.ShapeTree.AppendChild(new GroupShapeProperties());
            // Create the slide part for the new slide.
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();

            // Save the new slide part.
            slide.Save(slidePart);

            // Modify the slide ID list in the presentation part.
            // The slide ID list should not be null.
            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

            // Find the highest slide ID in the current list.
            uint maxSlideId = 1;
            SlideId prevSlideId = null;

            foreach (SlideId slideId in slideIdList.ChildElements)
            {
                if (slideId.Id > maxSlideId)
                {
                    maxSlideId = slideId.Id;
                }

                position--;
                if (position == 0)
                {
                    prevSlideId = slideId;
                }

            }

            maxSlideId++;

            // Get the ID of the previous slide.
            SlidePart lastSlidePart;

            if (prevSlideId != null)
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId);
            }
            else
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(((SlideId)(slideIdList.ChildElements[0])).RelationshipId);
            }

            // Use the same slide layout as that of the previous slide.
            if (null != lastSlidePart.SlideLayoutPart)
            {
                slidePart.AddPart(lastSlidePart.SlideLayoutPart);
            }

            // Insert the new slide into the slide list after the previous slide.
            SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
            newSlideId.Id = maxSlideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);
            string str = new StreamReader(slidePart.GetStream()).ReadToEnd();
            return slidePart.Slide;

            // Save the modified presentation.
            //presentationPart.Presentation.Save();

        }
    }
}