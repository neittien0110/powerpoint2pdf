using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace powerpoint2pdf
{
    class PowerpointFile
    {

        /// <summary>
        ///      Powerpoint Application Handler, run and control a Powerpoint process in background
        /// </summary>
        static Microsoft.Office.Interop.PowerPoint.Application ppApp = null;

        /// <summary>
        ///      Powerpoint slide Handler, run and control a Powerpoint process in background
        /// </summary>
        static Microsoft.Office.Interop.PowerPoint.Presentation objPres = null;

        static public int OpenPowerpointFile(String FileName)
        {
            try
            {
                // Open Powerpoint App
                ppApp = new Microsoft.Office.Interop.PowerPoint.Application();

                // Show the Powerpoint App on the screen (or may be not)
                ppApp.Visible = MsoTriState.msoTrue;


                Presentations presprint = ppApp.Presentations;

                // Slide Handler with a pptx file
                Presentation objPres = presprint.Open(FileName, MsoTriState.msoTrue,MsoTriState.msoTrue, MsoTriState.msoTrue);

                string PrinterName = ppApp.ActivePrinter;

                objPres.ExportAsFixedFormat(Path.ChangeExtension(FileName, ".pdf"),
                                PpFixedFormatType.ppFixedFormatTypePDF,
                                PpFixedFormatIntent.ppFixedFormatIntentPrint,
                                MsoTriState.msoFalse,
                                PpPrintHandoutOrder.ppPrintHandoutHorizontalFirst,
                                PpPrintOutputType.ppPrintOutputSlides,
                                MsoTriState.msoFalse,
                                null,
                                PpPrintRangeType.ppPrintAll,
                                "",
                                false,
                                false,
                                false,
                                true,
                                true,
                                System.Reflection.Missing.Value);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            finally

            {
                if (objPres !=null)
                {
                    objPres.Close();
                    objPres = null;
                }
                //Close the presentation without saving changes and quit PowerPoint
                if (ppApp !=null)
                {
                    ppApp.Quit();
                    ppApp = null;
                }
            }
            

            

            return 0;
        }
    }
}
