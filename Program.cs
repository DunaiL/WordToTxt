using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace WordToTxt
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string sourceFile = @"C:\source\file.docx";
            string targetFile = @"C:\target\file.txt";

            object missing = Type.Missing;
            object source = sourceFile;
            object target = targetFile;

            Application application = new Application();
            application.Visible = false;   // Не показывать Word юзеру  
            Document document = application.Documents.Open(ref source, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);  // Open the document  
            document.Activate();   // Make it the active document  

            object format = WdSaveFormat.wdFormatText;    // Save document as plain text  

            try
            {

                document.SaveAs(ref target, ref format);    // Save the document as a .txt file  

            }
            catch (Exception e)
            {

                Console.WriteLine("Exception: " + e);       // Handle any exceptions that may have occurred while saving the file  

            }
            finally
            {

                document.Close(ref missing);                 // Close the document and quit Word  

                application.Quit(ref missing);               // Release all resources used by Word application instance    

            }
        }
    }
}
