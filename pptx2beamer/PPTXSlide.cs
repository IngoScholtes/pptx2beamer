using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Xsl;
using System.Xml.XPath;

namespace pptx2beamer
{
    public class PPTXSlide
    {
        public int SlideNum;
        public XmlDocument SlideXml;
        public XmlDocument NotesXml;
        public XmlDocument SlideRelsXml;
        public XmlDocument NoteRelsXml;

        private XslCompiledTransform transform;

        XmlNamespaceManager ns;

        public PPTXSlide(int slideNum, XmlDocument slide, XmlDocument slideRel, XmlDocument notes, XmlDocument noteRels)
        {
            SlideNum = slideNum;
            SlideXml = slide;
            NotesXml = notes;
            SlideRelsXml = slideRel;
            NoteRelsXml = noteRels;

            ns = new XmlNamespaceManager(SlideXml.NameTable);
            ns.AddNamespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            transform = new System.Xml.Xsl.XslCompiledTransform();
            transform.Load("SlideTransformation.xslt");
        }

        public string LaTeX
        {
            get 
            {
                StringBuilder sb = new StringBuilder();
                XmlWriterSettings xmls = new XmlWriterSettings();
                xmls.ConformanceLevel = ConformanceLevel.Fragment;
                XPathDocument doc = new XPathDocument(new XmlNodeReader(SlideXml));
                XmlWriter xmlw = XmlWriter.Create(sb, xmls);
                transform.Transform(doc, xmlw);
                return sb.ToString();
            }
        }

        public bool HasTitle
        {
            get
            {
                string xpath = "//p:ph[@type='title']";
                return SlideXml.SelectSingleNode(xpath, ns) != null;
            }
        }

        public bool HasNotes
        {
            get 
            {
                return NotesXml != null;
            }
        }

        public string Notes
        {
            get 
            {
                return "";
            }
        }



        public bool IsTitleSlide
        {
            get 
            {
                string xpath = "//p:ph[@type='ctrTitle']";
                return SlideXml.SelectSingleNode(xpath, ns) != null;
            }
        }

        public string PresentationTitle
        {
            get 
            {
                string xpath = "//p:ph[@type='ctrTitle']/../../../p:txBody";
                return SlideXml.SelectSingleNode(xpath, ns).InnerText;
            }
        }

        public string SlideTitle
        {
            get
            {
                string xpath = "//p:ph[@type='title']/../../../p:txBody";                
                return SlideXml.SelectSingleNode(xpath, ns).InnerText;
            }
        }

        public bool HasText
        {
            get
            {
                string xpath = "//p:cNvPr[starts-with(@name, 'Inhaltsplatzhalter')]";
                return SlideXml.SelectSingleNode(xpath, ns) != null;
            }
        }

        public string SlideText
        {
            get
            {
                string xpath = "//p:cNvPr[starts-with(@name, 'Inhaltsplatzhalter')]/../../p:txBody";
                return SlideXml.SelectSingleNode(xpath, ns).InnerText;
            }
        }

        public bool HasImages
        {
            get
            {
                string xpath = "//p:pic";
                return SlideXml.SelectSingleNode(xpath, ns) != null;
            }
        }
    }
}
