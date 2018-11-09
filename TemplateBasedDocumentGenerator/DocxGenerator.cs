using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Xceed.Words.NET;

namespace TemplateBasedDocumentGenerator
{
    public class DocxGenerator
    {
        public static void FillTemplate(in Stream docxInputFileStream, Stream docxResultWriteStream, Dictionary<string, string> toReplace)
        {
            DocX doc = DocX.Load(docxInputFileStream);
            foreach (KeyValuePair<string, string> replace in toReplace)
            {
                doc.ReplaceText(replace.Key, replace.Value);
            }
            
            doc.SaveAs(docxResultWriteStream);
        }
    }
}
