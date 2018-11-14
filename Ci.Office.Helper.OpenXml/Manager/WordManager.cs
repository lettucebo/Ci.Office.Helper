using System;
using Ci.Extension;

namespace Ci.Office.Helper.OpenXml.Manager
{
    using System.Collections.Generic;
    using System.Drawing;
    using System.IO;
    using System.Linq;
    using System.Reflection;

    using Ci.Office.Helper.OpenXml.Models;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Drawing.Wordprocessing;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;

    using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
    using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
    using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

    public class WordManager
    {
        private static Assembly assembly;
        private static Stream imageStream;

        /// <summary>
        /// Contains the word processing document
        /// </summary>
        private WordprocessingDocument wordProcessingDocument;

        /// <summary>
        /// Contains the main document part
        /// </summary>
        private MainDocumentPart mainDocPart;

        /// <summary>
        /// Open an Word XML document 
        /// </summary>
        /// <param name="docname">name of the document to be opened</param>
        public void OpenDocuemnt(string docname)
        {
            // open the word docx
            wordProcessingDocument = WordprocessingDocument.Open(docname, true);

            // get the Main Document part
            mainDocPart = wordProcessingDocument.MainDocumentPart;
        }

        /// <summary>
        /// Close the document
        /// </summary>
        public void CloseDocument()
        {
            wordProcessingDocument.Close();
        }

        /// <summary>
        /// Updated text placeholders with texts.
        /// </summary>
        /// <param name="tagValueDict">Pair of placeholder tagID and text to replace.</param>
        public void UpdateText(Dictionary<string, OpenXmlTextInfo> tagValueDict)
        {
            foreach (var pair in tagValueDict)
            {
                var tagId = pair.Key;
                var text = pair.Value.Text;
                var isInnerXml = pair.Value.IsInnerXml;
                var customRunFont = pair.Value.RunFonts;

                foreach (var sdtElement in mainDocPart.Document.Body.Descendants<SdtElement>())
                {
                    if (sdtElement.SdtProperties.GetFirstChild<Tag>().Val == tagId)
                    {
                        OpenXmlElement parantElement = sdtElement.Descendants<Paragraph>().SingleOrDefault();
                        if (parantElement == null)
                        {
                            SdtContentRun cr = sdtElement.Descendants<SdtContentRun>().SingleOrDefault();
                            parantElement = cr;
                        }

                        if (parantElement != null)
                        {
                            Run r = parantElement.Descendants<Run>().FirstOrDefault();

                            if (r != null)
                            {
                                Text t = r.Descendants<Text>().SingleOrDefault();
                                if (t != null)
                                {
                                    if (isInnerXml)
                                    {
                                        r.InnerXml = text;
                                    }
                                    else
                                    {
                                        RunProperties runProperties = new RunProperties();

                                        // 判斷是否有自訂字元，不可以共用同一個 RunFonts 實體
                                        var runFont = new RunFonts();
                                        if (customRunFont == null)
                                        {
                                            runFont.Ascii = "Times New Roman";
                                            runFont.EastAsia = "標楷體";
                                        }
                                        else
                                        {
                                            runFont.Ascii = customRunFont.Ascii;
                                            runFont.EastAsia = customRunFont.EastAsia;
                                        }

                                        runProperties.Color = new DocumentFormat.OpenXml.Wordprocessing.Color()
                                        {
                                            Val = "000000"
                                        };
                                        runProperties.Append(runFont);

                                        // 判斷是否有換行字元，轉換成 break
                                        string[] stringSeparators = new string[] { "\r\n", "\n" };
                                        var textArr = text.Split(stringSeparators, StringSplitOptions.None).ToList();
                                        if (textArr.Last().IsNullOrWhiteSpace() && textArr.Count > 1)
                                        {
                                            // 判斷是否多行時，最後一行因 split 而多餘空白，將其移除
                                            textArr.RemoveAt(textArr.Count - 1);
                                        }

                                        foreach (string textValue in textArr)
                                        {
                                            runProperties.Append(new Text(textValue));
                                            if (textValue != textArr.Last())
                                            {
                                                runProperties.Append(new Break());

                                            }
                                        }

                                        r.AppendChild(runProperties);
                                        r.RemoveChild(t);
                                    }
                                }
                            }

                            break;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 動態產生表格
        /// </summary>
        public void UpdateTable(Dictionary<string, Table> tagValueDict)
        {
            Document dc = mainDocPart.Document;

            foreach (var pair in tagValueDict)
            {
                var tagId = pair.Key;
                var value = pair.Value;

                foreach (SdtBlock sb in dc.Descendants<SdtBlock>())
                {
                    //抓取Block 的名稱 產生要的表格
                    Tag tg = sb.Descendants<Tag>().FirstOrDefault(T => T.Val == tagId);
                    if (tg != null)
                    {
                        SdtContentBlock scb = sb.SdtContentBlock;
                        scb.ClearAllAttributes();
                        scb.Append(value);
                    }
                }
            }

        }

        /// <summary>
        /// Get the relationship id of image.
        /// </summary>
        /// <typeparam name="TSdtType">SdtElement type</typeparam>
        /// <param name="sdt">A sdtElement object that may contains image placeholder.</param>
        /// <param name="imageTag">Image placeholder tagID.</param>
        /// <returns>The relationship id of image.</returns>
        internal static string GetImageRelId<TSdtType>(TSdtType sdt, string imageTag) where TSdtType : SdtElement
        {
            // loop through all tags in the document within the sdt element
            foreach (Tag t in sdt.Descendants<Tag>().ToList())
            {
                // Do we have the correct tag?
                if (t.Val.ToString().ToUpper() == imageTag.ToUpper())
                {

                    // Get the BLIP for the image - there is only one image per placeholder so no need to loop through anything
                    DocumentFormat.OpenXml.Drawing.Blip b = sdt.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
                    if (null != b)
                    {
                        // return the image id tag
                        return b.Embed.Value;
                    }
                }
            }

            return string.Empty;
        }

        /// <summary>
        /// Get the original size of placeholder image.
        /// </summary>
        /// <param name="drawingList">Drawing object that may contains the image relationship id.</param>
        /// <param name="relId">The image relationship id.</param>
        /// <param name="width">Width of the image.</param>
        /// <param name="height">Height of the image.</param>
        internal static void GetPlaceholderImageSize(IEnumerable<Drawing> drawingList, string relId, out int width, out int height)
        {
            width = -1;
            height = -1;

            // Loop through all Drawing elements in the document
            foreach (Drawing d in drawingList)
            {
                // Loop through all the pictures (Blip) in the document
                if (d.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().ToList().Any(b => b.Embed.ToString() == relId))
                {
                    // The document size is in EMU. 1 pixel = 9525 EMU

                    // The size of the image placeholder is located in the EXTENT element
                    Extent e = d.Descendants<Extent>().FirstOrDefault();
                    if (null != e)
                    {
                        width = (int)(e.Cx / 9525);
                        height = (int)(e.Cy / 9525);
                    }

                    if (width == -1)
                    {
                        // The size of the image is located in the EXTENTS element
                        DocumentFormat.OpenXml.Drawing.Extents e2 = d.Descendants<DocumentFormat.OpenXml.Drawing.Extents>().FirstOrDefault();
                        if (null != e2)
                        {
                            width = (int)(e2.Cx / 9525);
                            height = (int)(e2.Cy / 9525);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Replace the image part with image memory stream.
        /// </summary>
        /// <param name="relId">The relationship id of the placeholder image.</param>
        /// <param name="imageStream">Image memory stream to replace the placeholder image.</param>
        /// <param name="width">Width of placeholder image.</param>
        /// <param name="height">Height of placeholder image.</param>
        private void UpdateImagePart(string relId, MemoryStream imageStream, int width, int height)
        {
            var originalBitmap = Image.FromStream(imageStream);
            var bitmap = originalBitmap;

            // resize image
            if (width != -1)
            {
                bitmap = new Bitmap(originalBitmap, width, height);
            }

            // Save image data to ImagePart
            var stream = new MemoryStream();
            bitmap.Save(stream, originalBitmap.RawFormat);

            // Get the ImagePart
            var imagePart = (ImagePart)mainDocPart.GetPartById(relId);

            // Create a writer to the ImagePart
            var writer = new BinaryWriter(imagePart.GetStream());

            // Overwrite the current image in the docx file package
            writer.Write(stream.ToArray());

            // Close the ImagePart
            writer.Close();
        }

        /// <summary>
        /// Updated image placeholders with images.
        /// </summary>
        /// <param name="tagValueDict">Pair of placeholder tagID and image to replace.</param>
        public void UpdateImage(Dictionary<string, MemoryStream> tagValueDict)
        {
            foreach (var pair in tagValueDict)
            {
                var tagId = pair.Key;
                var imageStream = pair.Value;

                foreach (SdtElement sdtElement in mainDocPart.Document.Body.Descendants<SdtElement>())
                {
                    string relId = GetImageRelId(sdtElement, tagId);
                    if (!string.IsNullOrEmpty(relId))
                    {
                        if (imageStream == null)
                        {
                            sdtElement.Remove();
                            continue;
                        }

                        // Get size of image
                        int imageWidth;
                        int imageHeight;
                        GetPlaceholderImageSize(mainDocPart.Document.Body.Descendants<Drawing>(), relId, out imageWidth, out imageHeight);

                        UpdateImagePart(relId, imageStream, imageWidth, imageHeight);

                        break;
                    }
                }
            }
        }

        /// <summary>
        /// Image file path to MemoryStream
        /// </summary>
        /// <param name="imagePath">The image path.</param>
        /// <returns>MemoryStream.</returns>
        public static MemoryStream GetStreamFromImagePath(string imagePath)
        {
            if (!File.Exists(imagePath))
            {
                // can not find image, use default
                // read embedded resource
                assembly = Assembly.GetExecutingAssembly();
                imageStream = assembly.GetManifestResourceStream("Ci.Office.Helper.OpenXml.Resources.Default.png");

                MemoryStream ms = new MemoryStream();
                imageStream.CopyTo(ms);
                return ms;
            }
            else
            {
                byte[] original = File.ReadAllBytes(imagePath);
                var stream = new MemoryStream(original);
                return stream;
            }
        }

    }
}

