using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ED47.WordTemplate
{
    public class WordTemplateData : ObservableCollection<IElementData>
    {
        private IEnumerable<SdtElement> _template;

        public WordTemplateData()
        {
            CollectionChanged += (sender, args) =>
            {
                if (args.Action != NotifyCollectionChangedAction.Add)
                    return;

                foreach (IElementData newItem in args.NewItems)
                {
                    newItem.Template = this.Template;
                }
            };
        }

        public IEnumerable<SdtElement> Template
        {
            get { return _template; }
            set
            {
                _template = value;

                foreach (var item in this)
                {
                    item.Template = Template;
                }
            }
        }

        public void Apply(IEnumerable<OpenXmlElement> elements, bool isTopLevel)
        {
            var sdts = elements
                .SelectMany(el => el.Descendants<SdtElement>())
                .ToList();

            sdts.AddRange(elements.OfType<SdtElement>());

            foreach (var data in this)
            {
                var matchingTags = sdts
                    .Where(el => OpenXmlHelper.GetTag(el) == data.TagName)
                    .ToList();

                if (!matchingTags.Any())
                    continue;

                data.Apply(matchingTags, isTopLevel);
            }
        }

        public void Apply(WordprocessingDocument wordDocument)
        {
            foreach (var headerPart in wordDocument.MainDocumentPart.HeaderParts)
            {
                Apply(headerPart.RootElement, isTopLevel: true);
            }

            Apply(wordDocument.MainDocumentPart.RootElement, isTopLevel: true);

            foreach (var footerPart in wordDocument.MainDocumentPart.FooterParts)
            {
                Apply(footerPart.RootElement, isTopLevel: true);
            }
        }
    }

    public class FieldData : IElementData
    {
        public string TagName { get; set; }
        public string Value { get; set; }

        public IEnumerable<SdtElement> Template { get; set; }

        public void Apply(IEnumerable<SdtElement> sdts, bool isTopLevel)
        {
            foreach (var sdt in sdts)
            {
                if (isTopLevel)
                {
                    //Maintain SDTs not in a collection
                    var text = sdt.Descendants<Text>().FirstOrDefault();

                    if (text == null)
                        continue;

                    text.Text = Value;
                }
                else
                {
                    var run = sdt.Descendants<Run>().FirstOrDefault();

                    if (run == null)
                        continue;

                    run = (Run)run.CloneNode(true);

                    var text = run.Descendants<Text>().FirstOrDefault();

                    if (text == null)
                        continue;

                    text.Text = Value;
                    sdt.Parent.InsertAfter(run, sdt);
                    sdt.Remove();
                }
            }
        }
    }

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

    public interface IElementData
    {
        string TagName { get; }
        void Apply(IEnumerable<SdtElement> sdts, bool isTopLevel);
        IEnumerable<SdtElement> Template { get; set; }
    }

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