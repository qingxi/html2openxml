using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using HtmlToOpenXml;
using DocumentFormat.OpenXml.Wordprocessing;


namespace ConsoleApp1
{
    class Program
    {
        private static readonly int word_max_width = 560;
        private static readonly int word_max_height = 860;
        static void Main(string[] args)
        {

            string html = File.ReadAllText(@"C:\Users\shisx\Source\Repos\html2openxml\test1.html");
            using (MemoryStream generatedDocument = new MemoryStream())
            {
                using (WordprocessingDocument package = WordprocessingDocument.Create(generatedDocument, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = package.MainDocumentPart;
                    if (mainPart == null)
                    {
                        mainPart = package.AddMainDocumentPart();
                        new Document(new Body()).Save(mainPart);
                    }
                    HtmlConverter converter = new HtmlConverter(mainPart);
                    converter.ImageProcessing = ImageProcessing.Base64Provisioning;
                    converter.ProvisionImage += (sender, e) =>
                    {
                        if (e.ImageSize.IsEmpty || e.ImageSize.Width > word_max_width || e.ImageSize.Height > word_max_height)
                        {
                            var width = e.ImageSize.Width * 1.0;
                            var height = e.ImageSize.Height * 1.0;
                            if (width > word_max_width)
                            {
                                width = word_max_width;
                                height = (word_max_width * 1.00 / e.ImageSize.Width) * e.ImageSize.Height;
                            }

                            if (height > word_max_height)
                            {
                                height = word_max_height;
                                width = (word_max_height * 1.00 / e.ImageSize.Height) * e.ImageSize.Width;
                            }

                            e.ImageSize = new Size((int)width, (int)height);
                        }
                    };
                    converter.ConsiderDivAsParagraph = true;
                    converter.RenderPreAsTable = true;
                    Body body = mainPart.Document.Body;

                    var paragraphs = converter.Parse(html);
                    for (int i = 0; i < paragraphs.Count; i++)
                    {
                        body.Append(paragraphs[i]);
                    }
                    mainPart.Document.Save();
                }
                File.WriteAllBytes(@"C:\Users\shisx\Source\Repos\html2openxml\test1.docx", generatedDocument.ToArray());
            }
        }
    }
}
