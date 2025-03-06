using System;
using System.IO;
using System.Collections.Generic;
using System.Runtime.Serialization.Formatters.Binary;

namespace ExcelTemplateManager
{
    public class TemplateManager
    {
        public void SaveTemplate(Template template, string filePath)
        {
            using (FileStream fileStream = new FileStream(filePath, FileMode.Create))
            {
#pragma warning disable SYSLIB0011
                BinaryFormatter formatter = new BinaryFormatter();
                formatter.Serialize(fileStream, template);
#pragma warning restore SYSLIB0011
            }
        }

        public Template LoadTemplate(string filePath)
        {
            using (FileStream fileStream = new FileStream(filePath, FileMode.Open))
            {
#pragma warning disable SYSLIB0011
                BinaryFormatter formatter = new BinaryFormatter();
                return (Template)formatter.Deserialize(fileStream);
#pragma warning restore SYSLIB0011
            }
        }

        public void SaveTemplatesToFile(Dictionary<string, Template> templates, string filePath)
        {
            using (FileStream fileStream = new FileStream(filePath, FileMode.Create))
            {
#pragma warning disable SYSLIB0011
                BinaryFormatter formatter = new BinaryFormatter();
                formatter.Serialize(fileStream, templates);
#pragma warning restore SYSLIB0011
            }
        }

        public Dictionary<string, Template> LoadTemplatesFromFile(string filePath)
        {
            using (FileStream fileStream = new FileStream(filePath, FileMode.Open))
            {
#pragma warning disable SYSLIB0011
                BinaryFormatter formatter = new BinaryFormatter();
                return (Dictionary<string, Template>)formatter.Deserialize(fileStream);
#pragma warning restore SYSLIB0011
            }
        }
    }
}