using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ED47.WordTemplate
{
    public class CollectionData : IElementData
    {
        private IEnumerable<SdtElement> _template;

        public string TagName { get; set; }
        
        public ICollection<WordTemplateData> ElementData { get; set; }

        public IEnumerable<SdtElement> Template
        {
            get { return _template; }
            set
            {
                _template = value;

                foreach (var item in this.ElementData)
                {
                    item.Template = Template;
                }
            }
        }

        public void Apply(IEnumerable<SdtElement> sdts, bool isTopLevel)
        {
            foreach (var sdt in sdts)
            {
                var tagName = OpenXmlHelper.GetTag(sdt);

                var templateSdt = Template.FirstOrDefault(el => OpenXmlHelper.GetTag(el) == tagName);

                if(templateSdt == null)
                    continue;
                
                var sdtContent = sdt.GetFirstChild<SdtContentBlock>();
                sdtContent.RemoveAllChildren();

                var templateElements = templateSdt.GetFirstChild<SdtContentBlock>().Elements().ToList();

                foreach (var elementData in ElementData)
                {
                    var allClones = new List<OpenXmlElement>();

                    foreach (var templateElement in templateElements)
                    {
                        var clone = templateElement.CloneNode(true);
                        allClones.Add(clone);

                        if (isTopLevel)
                            sdtContent.AppendChild(clone);
                        else
                        {
                            sdt.Parent.InsertAfter(clone, sdt);
                        }
                    }

                    if (!isTopLevel)
                        sdt.Remove();

                    elementData.Apply(allClones, isTopLevel: false);
                }
            }
        }
    }
}