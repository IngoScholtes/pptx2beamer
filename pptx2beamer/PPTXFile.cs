using Ionic.Zip;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace pptx2beamer
{
    public class PPTXFile
    {
        Dictionary<int, XmlDocument> slides = new Dictionary<int, XmlDocument>();
        Dictionary<int, XmlDocument> slide_rels = new Dictionary<int, XmlDocument>();
        Dictionary<int, XmlDocument> notes = new Dictionary<int, XmlDocument>();
        Dictionary<int, XmlDocument> note_rels = new Dictionary<int, XmlDocument>();

        string path;
        ZipFile pptx;


        public PPTXFile(string path)
        { 
            this.path = path;
            this.pptx = new ZipFile(path);
            Parse();
        }


        public int SlideCount
        {
            get
            {
                return slides.Count;
            }
        }

        public PPTXSlide TitleSlide
        {
            get
            { 
                return (from x in Slides where x.IsTitleSlide select x).First();
            }
        }

        public IEnumerable<PPTXSlide> Slides
        {
            get
            {
                List<PPTXSlide>  slides = new List<PPTXSlide>();
                for(int i=1; i<=SlideCount; i++)
                    slides.Add(GetSlide(i));
                return slides;
            }
        }

        public PPTXSlide GetSlide(int slide_num)
        {
            return new PPTXSlide(slide_num, 
                slides.ContainsKey(slide_num)?slides[slide_num]:null,
                slide_rels.ContainsKey(slide_num)?slide_rels[slide_num]:null,
                notes.ContainsKey(slide_num)?notes[slide_num]:null,
                note_rels.ContainsKey(slide_num)?note_rels[slide_num]:null);
        }

        public void ExtractMedia(string output_path)
        {
            if (!output_path.EndsWith("\\"))
                output_path = output_path + "\\";

            if (!System.IO.Directory.Exists(output_path))
                System.IO.Directory.CreateDirectory(output_path);

            foreach (ZipEntry entry in pptx.EntriesSorted)
            {
                // Extract media files
                if (entry.FileName.StartsWith("ppt/media/") && !entry.IsDirectory)
                {                    
                    System.IO.FileStream fs = new System.IO.FileStream(output_path + entry.FileName.Replace("ppt/media/", ""), System.IO.FileMode.Create);
                    entry.Extract(fs);
                }
            }        
        }

        private void Parse()
        {
            foreach (ZipEntry entry in pptx.EntriesSorted)
            {
                // extract slides
                if (entry.FileName.StartsWith("ppt/slides/slide") && !entry.IsDirectory)
                {
                    int slide_num = int.Parse(entry.FileName.Replace("ppt/slides/slide", "").Replace(".xml", ""));
                    System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                    doc.Load(entry.OpenReader());
                    slides[slide_num] = doc;
                }
                // extract links to media files
                else if (entry.FileName.StartsWith("ppt/slides/_rels") && !entry.IsDirectory)
                {
                    int slide_num = int.Parse(entry.FileName.Replace("ppt/slides/_rels/slide", "").Replace(".xml.rels", ""));
                    System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                    doc.Load(entry.OpenReader());
                    slide_rels[slide_num] = doc;
                }
                // extract slide notes
                else if (entry.FileName.StartsWith("ppt/notesSlides/notesSlide") && !entry.IsDirectory)
                {
                    int slide_num = int.Parse(entry.FileName.Replace("ppt/notesSlides/notesSlide", "").Replace(".xml", ""));
                    System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                    doc.Load(entry.OpenReader());
                    notes[slide_num] = doc;
                }
                // extract links to media files
                else if (entry.FileName.StartsWith("ppt/notesSlides/_rels") && !entry.IsDirectory)
                {
                    int slide_num = int.Parse(entry.FileName.Replace("ppt/notesSlides/_rels/notesSlide", "").Replace(".xml.rels", ""));
                    System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                    doc.Load(entry.OpenReader());
                    note_rels[slide_num] = doc;
                }
            }
        }
    }
}
