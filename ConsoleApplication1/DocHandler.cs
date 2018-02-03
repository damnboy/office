using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace office
{
    /*
     * https://msdn.microsoft.com/en-us/library/office/aa192495(v=office.11).aspx
     * characters -> words -> sentences -> paragraphs -> sections (headers & footers) 
     */


    class DocTextContentHandler
    {
        Word.Application _app;
        public DocTextContentHandler()
        {
            
        }

        public bool initApp()
        {
            if (_app == null)
            {
                _app = new Microsoft.Office.Interop.Word.Application();
                _app.Visible = false;
            }
            
            return _app != null;
        }

        public void closeApp()
        {
            if (_app != null)
            {
                foreach (Word.Document doc in _app.Documents)
                {
                    doc.Close();
                }
                _app.Quit();
                _app = null;
            }
        }

        public int process(String doc)
        {
            if (!initApp()){
                return -1;
            }
            int handle = 0;
            try
            {
                Word.Document _doc = this._app.Documents.Open(doc);

                handle = processTables(_doc);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            finally
            {
                this.closeApp();
            }
            return handle;
        }

        public int processTables(Word.Document doc)
        {
            int handle = 0;
            if (doc != null)
            {
                foreach (Word.Table table in doc.Tables)
                {
                    try{
                        if (this.onTable(doc, table))
                        {
                            handle++;
                        }
                    }
                    catch(Exception e){
                        Console.WriteLine(e);
                    }
                    
                }
            }
            return handle;
        }

        public virtual bool onTable(Word.Document doc, Word.Table t)
        {
            if (t.Columns.Count == 4)
            {
                //Console.Write(doc.Name + "\t[" + t.Rows.Count + "]\t");
                //IP\端口\服务
                for (int r = 2; r <= t.Rows.Count; r++)
                {
                    Console.WriteLine(doc.Name.Replace(".doc", "") + "\t" + t.Cell(r, 2).Range.Text.Replace("\r\a", "").Trim() + "\t" +
                        t.Cell(r, 1).Range.Text.Replace("\r\a", "").Trim() + "\t" +
                        t.Cell(r, 3).Range.Text.Replace("\r\a", "").Trim() + "\t");
                }
            }
            return true;
        }

    }
}

