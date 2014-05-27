using System;
using System.Collections.Generic;
using System.IO;

namespace ED47.WordTemplate.Demo
{
    class Program
    {
        static void Main(string[] args)
        {
            var data = new WordTemplateData
            {
                new FieldData
                {
                    TagName = "World",
                    Value = "people"
                },
                new CollectionData
                {
                    TagName = "Chapter",
                    ElementData = new List<WordTemplateData>
                    { 
                        new WordTemplateData
                        {
                            new FieldData
                            {
                                TagName = "Date",
                                Value = DateTime.Now.ToShortDateString()
                            },
                            new FieldData
                            {
                                TagName = "Name",
                                Value = "Name 1"
                            },
                            new CollectionData
                            {
                                TagName = "SubChapter",
                                ElementData = new List<WordTemplateData>
                                {
                                    new WordTemplateData
                                    {
                                        new FieldData
                                        {
                                            TagName = "Change",
                                            Value = Guid.NewGuid().ToString()
                                        }
                                    },
                                    new WordTemplateData
                                    {
                                        new FieldData
                                        {
                                            TagName = "Change",
                                            Value = Guid.NewGuid().ToString()
                                        }
                                    }
                                }
                            }
                        },
                        new WordTemplateData
                        {
                            new FieldData
                            {
                                TagName = "Date",
                                Value = DateTime.Now.ToShortDateString()
                            },
                            new FieldData
                            {
                                TagName = "Name",
                                Value = "Name 222"
                            },
                            new CollectionData
                            {
                                TagName = "SubChapter",
                                ElementData = new List<WordTemplateData>
                                {
                                    new WordTemplateData
                                    {
                                        new FieldData
                                        {
                                            TagName = "Change",
                                            Value = Guid.NewGuid().ToString()
                                        }
                                    }
                                }
                            }
                        },
                        new WordTemplateData
                        {
                            new FieldData
                            {
                                TagName = "Date",
                                Value = DateTime.Now.AddDays(30).ToShortDateString()
                            },
                            new FieldData
                            {
                                TagName = "Name",
                                Value = "Name 3"
                            },
                            new CollectionData
                            {
                                TagName = "SubChapter",
                                ElementData = new List<WordTemplateData>
                                {
                                    new WordTemplateData
                                    {
                                        new FieldData
                                        {
                                            TagName = "Change",
                                            Value = Guid.NewGuid().ToString()
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            };

            const string template = "InputTemplate.docx";
            const string output = "output.docx";
            MemoryStream outputStream = null;
            Stream existingStream = null;

            try
            {
                using (var templateStream = File.Open(template, FileMode.Open))
                {
                    if (File.Exists(output))
                        existingStream = File.Open(output, FileMode.Open);

                    outputStream = WordTemplate.Generate(data, templateStream, existingStream);
                }

                if (existingStream != null)
                    existingStream.Close();

                using (var outputWrite = new FileStream(output, FileMode.Create))
                {
                    outputStream.WriteTo(outputWrite);
                    outputStream.Close();
                }
            }
            finally
            {
                if (outputStream != null)
                    outputStream.Close();

                if (existingStream != null)
                    existingStream.Close();
            }
        }
    }
}
