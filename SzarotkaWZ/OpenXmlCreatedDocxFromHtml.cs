using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using SzarotkaWZ.Helper;
using SzarotkaWZ.Models;

namespace SzarotkaWZ
{
    internal class OpenXmlCreatedDocxFromHtml
    {

        public static void CreateHtml(string path, List<Wz> wzs)
        {
            // MainDocumentPart mainPart;
            Body b;
            Document d;
            AlternativeFormatImportPart chunk;
            AltChunk altChunk;

            string altChunkID = "AltChunkId1";

            using var ms = new MemoryStream();

            using (var myDoc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
            {
                // mainPart = myDoc.MainDocumentPart;

                // Add a main document part. 
                MainDocumentPart mainPart = myDoc.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());


                if (mainPart == null)
                {
                    mainPart = myDoc.AddMainDocumentPart();
                    b = new Body();
                    d = new Document(b);
                    d.Save(mainPart);
                }

                chunk = mainPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Xhtml, altChunkID);

                using (Stream chunkStream = chunk.GetStream(FileMode.Create, FileAccess.Write))
                {
                    using (StreamWriter stringStream = new StreamWriter(chunkStream, System.Text.Encoding.UTF8))
                    {
                        string html = HtmlTable.StartHtml;
                        for (int i = 0; i < wzs.Count; i++)
                        {
                            html += HtmlTable.Created(wzs[i]);
                            html += "<br>";
                        }
                        html += HtmlTable.EndHtml;
                        stringStream.Write(html);
                    }
                }

                altChunk = new AltChunk();
                altChunk.Id = altChunkID;
                mainPart.Document.Body.InsertAt<AltChunk>(altChunk, 1);
                mainPart.Document.Save();
            }
        }

    }


}

