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

        private XslCompiledTransform slideTransform;
        private XslCompiledTransform titleTransform;
        private XslCompiledTransform notesTransform;

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

            slideTransform = new System.Xml.Xsl.XslCompiledTransform();
            slideTransform.Load("SlideTransformation.xslt");

            titleTransform = new System.Xml.Xsl.XslCompiledTransform();
            titleTransform.Load("TitleTransformation.xslt");

            notesTransform = new System.Xml.Xsl.XslCompiledTransform();
            notesTransform.Load("NotesTransformation.xslt");
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

                if (IsTitleSlide)
                    titleTransform.Transform(doc, xmlw);
                else
                    slideTransform.Transform(doc, xmlw);

                if (HasNotes)
                    notesTransform.Transform(new XPathDocument(new XmlNodeReader(NotesXml)), xmlw);

                return sb.ToString();
            }
        }      

        public bool HasNotes
        {
            get 
            {
                return NotesXml != null;
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
    }
}
