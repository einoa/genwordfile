using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;


namespace GenWordFile1
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create Package
            System.IO.Packaging.Package pkg;
            pkg = Package.Open(@"C:\TMP\z.docx", FileMode.Create, FileAccess.ReadWrite);


            string nsWordprocessingML = @"http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            XmlDocument doc = new XmlDocument();


            XmlElement xmlDocument = doc.CreateElement("w:document", nsWordprocessingML);
            doc.AppendChild(xmlDocument);

            XmlElement xmlBody = doc.CreateElement("w:body", nsWordprocessingML);
            xmlDocument.AppendChild(xmlBody);

            XmlElement xmlParagraph = doc.CreateElement("w:p", nsWordprocessingML);
            xmlBody.AppendChild(xmlParagraph);

            XmlElement xmlRun = doc.CreateElement("w:r", nsWordprocessingML);
            xmlParagraph.AppendChild(xmlRun);

            XmlElement xmlText = doc.CreateElement("w:t", nsWordprocessingML);
            xmlRun.AppendChild(xmlText);

            XmlNode nodeText = doc.CreateNode(XmlNodeType.Text, "w:t", nsWordprocessingML);
            nodeText.Value = "Hello World";
            xmlText.AppendChild(nodeText);


            // Write document.xml
            Uri uriDocument;
            uriDocument = new Uri("/word/document.xml", UriKind.Relative);

            PackagePart partDocument = pkg.CreatePart(uriDocument, "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml");
            StreamWriter swDocument = new StreamWriter(partDocument.GetStream(FileMode.Create, FileAccess.Write));
            doc.Save(swDocument);
            swDocument.Close();
            pkg.Flush();

            // Write relationships
            uriDocument = new Uri("/word/document.xml", UriKind.Relative);
            PackageRelationship rel = pkg.CreateRelationship(uriDocument, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "rId1");
            pkg.Flush();

            // Close
            pkg.Close();
        }
    }
}
