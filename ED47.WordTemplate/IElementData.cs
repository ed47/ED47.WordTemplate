using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ED47.WordTemplate
{
    public interface IElementData
    {
        string TagName { get; }
        void Apply(IEnumerable<SdtElement> sdts, bool isTopLevel);
        IEnumerable<SdtElement> Template { get; set; }
    }
}