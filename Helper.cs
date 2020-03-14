using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace WordManulipation
{

    public class Helper
    {

        public static byte[] WordManipulation(byte[] documentByte, byte[] imageByte)
        {
            var ms = new MemoryStream();
            ms.Write(documentByte, 0, documentByte.Length);

            using (WordprocessingDocument wordprocessingDocument =
                WordprocessingDocument.Open(ms, true))
            {
                MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;

                if (imageByte != null)
                {
                    var cat1Img = new ImageData("test", imageByte)
                    {
                        Width = new Decimal(10.81),
                        Height = new Decimal(3.51)
                    };

                    var imgRun = DocxImgHelper.GenerateImageRun(wordprocessingDocument, cat1Img);

                    //firstShapeAddOrReplaceText(mainPart.Document.Body, PATIENT_LABEL_STRING, imgRun);
                    firstShapeAddOrReplaceImage(mainPart.Document.Body, imgRun);
                }

                removeDoctorSignatureString(mainPart.Document);
            }
            return ms.ToArray();
        }

        static Regex getDoctorSignatureRegex()
        {
            return new Regex(@"Doctor.*signature.*", RegexOptions.IgnoreCase);
        }
        static string getdoctorSignatureStringFromXml(string xml)
        {
            var match = getDoctorSignatureRegex().Match(xml);

            return match == null ? "" : match.Value;
        }
     
        static void firstShapeAddOrReplaceImage(Body body, Run imageRun)
        {
            OpenXmlElement shape = null;
      
            var wordProcessingShape = body.Descendants<Shape>().FirstOrDefault();
            if (wordProcessingShape != null)
            {
                shape = wordProcessingShape;
            }
            else //Aspose lookup for roundRect..
            {
                var roundRect = body.Descendants<RoundRectangle>().FirstOrDefault();
                if (roundRect != null)
                {
                    shape = roundRect;                    
                }
            }
            if (shape != null)
            {
                Run parentRun = null;
                OpenXmlElement tmp = shape.Parent;
                while (tmp.Parent != null && !(tmp is Run))
                {
                    tmp = tmp.Parent;
                }
                parentRun = tmp as Run;
                if (parentRun != null)
                {
                    parentRun.InnerXml = imageRun.InnerXml;
                }
            }


            Action<OpenXmlCompositeElement> findParentRunAndEmptyIt = (x) =>
            {
                Run parentRun = null;
                OpenXmlElement tmp = x.Parent;                
                while (tmp.Parent != null && !(tmp is Run))
                {
                    tmp = tmp.Parent;
                }
                parentRun = tmp as Run;
                if (parentRun != null)
                {
                    parentRun.InnerXml = "";
                }
            };

            //Remove both absolute 
            var l1 = body.Descendants<Shape>().ToList();
            if(l1.Count>0)
            {
                var isAbsolute = l1[0].Style.Value.Contains("position:absolute");
                if(isAbsolute)
                findParentRunAndEmptyIt(l1[0]);
            }
            var l2 = body.Descendants<RoundRectangle>().ToList();
            if(l2.Count>0)
            {


                var isAbsolute = l2[0].Style.Value.Contains("position:absolute");
                var val = l2[0].Style.Value.Split(';');
                Dictionary<string, string> dict = new Dictionary<string, string>();
                foreach(var v in val )
                {
                    var splited = v.Split(':');
                    var key = splited[0].Trim();
                    var value = splited[1].Trim();
                    if(!dict.ContainsKey(key))
                        dict.Add(key, value);
                }
                if (isAbsolute)
                {
                    findParentRunAndEmptyIt(l2[0]);
                }

            }
        }


        static void removeDoctorSignatureString(Document document)
        {
            var body = document.Body;

            Predicate<string> getDoctorSignatureXml = new Predicate<string>((xml) =>
            {
                var match = getDoctorSignatureRegex().Matches(xml);
                return match.Count > 0;
            });
            var paragraph = body.Descendants<Paragraph>();
            var targetControl = body.Descendants<Paragraph>().FirstOrDefault(x => getDoctorSignatureXml(x.InnerText));
            //In case not found
            if (targetControl == null)
            {
                return;
            }

            //Remove it
            targetControl.InnerXml = "";

        }

    }
}

