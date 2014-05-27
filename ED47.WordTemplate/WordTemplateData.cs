using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Linq;
using System.Reflection;
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

        public void Load(string collectionTagName, IEnumerable<object> data, string name = null)
        {
            var collectionData = new List<WordTemplateData>();

            foreach (var o in data)
            {
                var templateData = new WordTemplateData();
                templateData.Load(o, name);
                collectionData.Add(templateData);
            }

            Add(new CollectionData { TagName = collectionTagName, ElementData = collectionData });
        }

        public void Load(object data, string name = null)
        {
            if (data == null)
                return;

            var type = data.GetType();

            var properties = type.GetProperties(BindingFlags.Instance | BindingFlags.Public);

            foreach (var propertyInfo in properties)
            {
                var value = propertyInfo.GetValue(data);

                if (propertyInfo.PropertyType.IsGenericType && propertyInfo.PropertyType.GetInterface("IEnumerable") != null)
                {
                    if (value != null)
                    {
                        Load(String.Format("{0}.{1}", name ?? type.Name, propertyInfo.Name), (IEnumerable<object>) value);
                    }

                    continue;
                }
                
                Add(new FieldData
                {
                    TagName = String.Format("{0}.{1}", name ?? type.Name, propertyInfo.Name),
                    Value = value != null ? value.ToString().Trim() : String.Empty
                });
            }
        }

        public void Load(string name, string data)
        {
            if (data == null)
                return;
            
            Add(new FieldData
            {
                TagName = name,
                Value = data
            });
        }
    }
}