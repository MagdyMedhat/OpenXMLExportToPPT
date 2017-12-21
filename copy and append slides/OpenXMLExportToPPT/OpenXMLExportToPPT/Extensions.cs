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
            UInt32 slideId = 256U;
            slideId += Convert.ToUInt32(presentationPart.Presentation.SlideIdList.Count());

            Slide slide = new Slide(new CommonSlideData(new ShapeTree()));

            SlidePart sPart = presentationPart.AddNewPart<SlidePart>();
            slide.Save(sPart);

            SlideMasterPart smPart = presentationPart.SlideMasterParts.First();
            SlideLayoutPart slPart = smPart.SlideLayoutParts.SingleOrDefault
                (sl => sl.SlideLayout.CommonSlideData.Name.Value.Equals(layoutName));
            if (slPart == null)
                throw new Exception("The slide layout " + layoutName + " is not found");
            sPart.AddPart<SlideLayoutPart>(slPart);

            sPart.Slide.CommonSlideData = (CommonSlideData)smPart.SlideLayoutParts.SingleOrDefault(
                sl => sl.SlideLayout.CommonSlideData.Name.Value.Equals(layoutName)).SlideLayout.CommonSlideData.Clone();

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
