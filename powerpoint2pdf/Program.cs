using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace powerpoint2pdf
{
    class Program
    {
        static void Main(string[] args)
        {
            ///@brief  Name of PPTX File which need to be convert to pdf.
            string FileName;

            /// Get demo pptx file
            FileName =  Directory.GetCurrentDirectory() + @"\demo.pptx";

            /// Convert
            PowerpointFile.OpenPowerpointFile(FileName);
        }
    }
}
