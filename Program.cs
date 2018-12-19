using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var doc = WordprocessingDocument.Create(@"D:\myfile.docx", WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                run.AppendChild(new Text("this is a new document"));
            }
        }
    }
}
/*
Create a new CS Console Application.
Add Reference > DocumentFormat.OpenXml.dll (Browse)
Add Reference > WindowsBase (Assembly>Framework)
*/
