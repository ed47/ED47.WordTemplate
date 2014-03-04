using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ED47.WordTemplate
{
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
}