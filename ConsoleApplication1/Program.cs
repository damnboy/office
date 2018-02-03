using System;  
using System.Collections.Generic;  
using System.Linq;  
using System.Text;  
using Word = Microsoft.Office.Interop.Word;

namespace office
{
    class Program
    {
        static void Main(string[] args)
        {
            DocTextContentHandler handler = new DocTextContentHandler();
            try
            {
                handler.process(@"C:\Users\Administrator\Desktop\safe.doc");
            }
            catch (Exception ex)
            {                
                Console.WriteLine(ex.Message);
            }
            finally{
                handler.closeApp();
            }
        }
    }
}