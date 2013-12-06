using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ionic.Zip;
using System.Xml;

namespace pptx2beamer
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length<1)
            {
                Console.WriteLine("Usage: pptx2beamer [pptxfile]");
                return;
            }

            PPTXFile pptx = new PPTXFile(args[0]);

            pptx.ExtractMedia("latex\\img");

            foreach (PPTXSlide slide in pptx.Slides)
            {
                if(slide.HasTitle)
                    Console.WriteLine(slide.Title);
                var texts = slide.SlideXml.GetElementsByTagName("p:sp");
                var images = slide.SlideXml.GetElementsByTagName("p:pic");
            }

        }

    }
}
