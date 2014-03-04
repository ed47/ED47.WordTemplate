using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ED47.WordTemplate
{
    public static class OpenXmlHelper
    {
        public static string GetTag(SdtElement sdt)
        {
            var sdtProperties = sdt.SdtProperties;

            if (sdtProperties == null || !sdtProperties.Any())
                return null;

            var tag = sdtProperties.Elements<Tag>().FirstOrDefault();

            if (tag == null)
                return null;

            return tag.Val;
        }
    }
}