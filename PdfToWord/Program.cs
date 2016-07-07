using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PdfToWord
{
    class Program
    {
        public static Microsoft.Office.Interop.Word.Document wordDocument { get; set; }
        static void Main(string[] args)
        {
            Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();
            try
            {
                var files = new DirectoryInfo(".").GetFiles();
                foreach (var item in files)
                {
                    try
                    {
                        wordDocument = appWord.Documents.Open(item.FullName);
                        wordDocument.SaveAs2(item.FullName.Replace(".pdf", ".docx"));
                    }
                    catch (Exception e)
                    {

                    }
                }
            }
            catch (Exception x)
            {
            }
            appWord.Quit();
            
        }
    }
}
