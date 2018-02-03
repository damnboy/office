using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace office
{    
    /* embedded objects
     * https://msdn.microsoft.com/en-us/library/office/hh965731(v=office.14).aspx
     * https://social.msdn.microsoft.com/Forums/vstudio/en-US/85ef2249-0344-42f5-8dec-e7c09f98c62b/extract-embedded-document-with-the-word-document?forum=vsto
     */
    
    class EmbeddedObjectHandler
    {
    }
}

/*
public virtual bool onOleObject()
{
    return true;
}
public int processEmbeddedOLEObjects(Word.Document doc)
{

    int handle = 0;
    if (doc != null)
    {
        foreach (Word.InlineShape inlineShape in doc.InlineShapes)
        {
            Word.WdInlineShapeType type = inlineShape.Type;
            if (type == Word.WdInlineShapeType.wdInlineShapeEmbeddedOLEObject)
            {
                Console.WriteLine(inlineShape.OLEFormat.IconLabel);
                Console.WriteLine(inlineShape.OLEFormat.ProgID);
            }

        }
    }
    return handle;
}
*/