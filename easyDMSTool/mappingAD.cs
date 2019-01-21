using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace easyDMSTool
{
    public class mappingAD
    {
        public mappingAD()
        {
        }
        public mappingAD(string groupAD, string countryCode, string department)
        {
            this.countryCode = countryCode;            
            this.groupAD = groupAD;
            this.department = department;
        }

        public void Build()
        {
        }


        static public void Serialize(System.Collections.Generic.List<mappingAD> button)
        {
            System.Xml.Serialization.XmlSerializer serializer = new System.Xml.Serialization.XmlSerializer(typeof(System.Collections.Generic.List<mappingAD>));
            System.IO.TextWriter writer = new System.IO.StreamWriter(@"C:\Users\Public\EasySender\dropdownxmlfilelist.xml");
            serializer.Serialize(writer, button);
            writer.Close();
        }

        static public List<mappingAD> Deserialize()
        {
            List<mappingAD> buttons;
            System.IO.StreamReader reader = new System.IO.StreamReader("\\\\DEIS335\\SendToEasyDMS\\config\\AD_mapping.xml");
            System.Xml.Serialization.XmlSerializer serializer = new System.Xml.Serialization.XmlSerializer(typeof(List<mappingAD>));
            buttons = (List<mappingAD>)serializer.Deserialize(reader);
            reader.Close();
            return buttons;
        }


        [System.Xml.Serialization.XmlAttribute("countryCode")]
        public string countryCode { get; set; }

        [System.Xml.Serialization.XmlAttribute("department")]
        public string department { get; set; }

        [System.Xml.Serialization.XmlElement("mapping")]
        public string groupAD { get; set; }

    }

}

