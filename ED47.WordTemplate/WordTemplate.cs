using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ED47.WordTemplate
{
    public static class WordTemplate
    {
        public static MemoryStream Generate(WordTemplateData data, Stream template, Stream existingDocument = null)
        {
            var output = new MemoryStream();

            if (existingDocument != null)
                existingDocument.CopyTo(output);
            else
                template.CopyTo(output);

            using (var templateDocument = WordprocessingDocument.Open(template, true))
            using (var wordDocument = WordprocessingDocument.Open(output, true))
            {
                data.Template = templateDocument.MainDocumentPart.RootElement.Descendants<SdtElement>();
                data.Apply(wordDocument);
            }

            return output;
        }
    }
}
