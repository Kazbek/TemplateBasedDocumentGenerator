using System.Collections.Generic;
using System.IO;
using Xceed.Words.NET;

namespace TemplateBasedDocumentGenerator
{
    public class DocxGenerator
    {
        public static void FillTemplate(in Stream docxTemplateFileStream, Stream docxResultWriteStream, in Dictionary<string, string> toReplace)
        {
            DocX doc = DocX.Load(docxTemplateFileStream);
            FillTemplate(doc, toReplace);
            
            doc.SaveAs(docxResultWriteStream);
        }

        public static void FillTemplate(in string docxTemplateFilePath, string docxResultFilePath, in Dictionary<string, string> toReplace)
        {
            DocX doc = DocX.Load(docxTemplateFilePath);
            FillTemplate(doc, toReplace);

            doc.SaveAs(docxResultFilePath);
        }

        public static void FillTemplate(in string docxTemplateFilePath, Stream docxResultWriteStream, in Dictionary<string, string> toReplace)
        {
            DocX doc = DocX.Load(docxTemplateFilePath);
            FillTemplate(doc, toReplace);

            doc.SaveAs(docxResultWriteStream);
        }

        public static void FillTemplate(in Stream docxTemplateFileStream, string docxResultFilePath, in Dictionary<string, string> toReplace)
        {
            DocX doc = DocX.Load(docxTemplateFileStream);
            FillTemplate(doc, toReplace);

            doc.SaveAs(docxResultFilePath);
        }

        private static void FillTemplate(DocX template, in Dictionary<string, string> toReplace)
        {
            foreach (KeyValuePair<string, string> replace in toReplace)
            {
                template.ReplaceText(replace.Key, replace.Value);
            }
        }
    }
}
