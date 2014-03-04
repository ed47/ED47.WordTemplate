using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ED47.WordTemplate
{
    public sealed class WordTemplateData : ObservableCollection<IElementData>
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
            var enumerable = elements as OpenXmlElement[] ?? elements.ToArray();

            var sdts = enumerable
                .SelectMany(el => el.Descendants<SdtElement>())
                .ToList();

            sdts.AddRange(enumerable.OfType<SdtElement>());

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
}