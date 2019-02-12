using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Jokedst.GetOpt;  //Sử dụng NuGet để cài đặt thêm gói Jokedst.GetOpt

namespace powerpoint2pdf
{
    class Program
    {
        static int Main(string[] args)
        {
            string InputDirectory = string.Empty;        /// input directory contains pptx files

            /// Use GetOpt function from the lib NuGet Jokedst.GetOpt to analyse commandline params
            try
            {
                var opts = new GetOpt("Convert PPTX to PDF", new[]
                {
                    new CommandLineOption('d', "--diretory", "input directory contains pptx files",  ParameterType.String, d => InputDirectory = (string)d)
                });

                opts.ParseOptions(args);

                if (InputDirectory==string.Empty)
                {
                    InputDirectory = Directory.GetCurrentDirectory();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return e.HResult;
            }

            //Open and scan all pptx files
            DirectoryInfo dir = new DirectoryInfo(InputDirectory);

            foreach (FileInfo f in dir.GetFiles("*.pptx"))
            {
                Console.WriteLine("File {0}", f.FullName);
            }

            return 0;
        }
    }
}
