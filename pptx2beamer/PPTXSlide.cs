using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
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
        }

        public bool HasTitle
        {
            get
            {
                string xpath = "//p:ph[@type='title']";
                return SlideXml.SelectSingleNode(xpath, ns) != null;
            }
        }

        public string Title
        {
            get
            {
                string xpath = "//p:ph[@type='title']/../../../p:txBody";                
                return SlideXml.SelectSingleNode(xpath, ns).InnerText;
            }
        }
    }
}
