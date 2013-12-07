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
            
            LaTeXExporter.ToLatex(pptx, "latex\\"+args[0].Replace(".pptx", ".tex"));
            Console.WriteLine("LaTeX written successfully.");
        }

    }
}
