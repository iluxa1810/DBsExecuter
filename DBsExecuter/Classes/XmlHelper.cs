using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace DBsExecuter.Classes
{
    [XmlRoot("Package")]
    public class Package
    {
        [XmlElement("Task")]
        public List<Task> Tasks { get; set; }
    }

    public class Task
    {
        [XmlAttribute("QueryName")]
        public string QueryName { get; set; }
        [XmlAttribute("QueryType")]
        public string QueryType { get; set; }
        [XmlText]
        public string Query { get; set; }
    }

    static class XmlHelper
    {
        public static Package GetXmlData(string path)
        {
            var xml = XDocument.Parse(File.ReadAllText(path));
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(Package));
            if (xmlSerializer.CanDeserialize(xml.CreateReader()))
            {
                Package serXml = (Package)xmlSerializer.Deserialize(xml.CreateReader());
                return serXml;
            }
            return null;
        }
    }
}


