using System;
using System.Xml.Serialization;

namespace ArmadaPack
{
    [Serializable]
    public class Vision
    {
        [XmlArray("TestParts")]
        public TestPart[] TestParts { get; set; }
    }
}
