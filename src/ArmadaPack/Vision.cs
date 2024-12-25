using System;
using System.Xml.Serialization;

namespace ArmadaPack
{
    [Serializable]
    public class Vision
    {
        public int PixelRatio { get; set; }
        public int BorderWidthPixels { get; set; }
        public int SimilarityRate { get; set; }
        public TestPartsGroup[] Groups { get; set; }
    }

    public class TestPartsGroup
    {
        [XmlAttribute]
        public int Site { get; set; }

        public TestPart[] TestParts { get; set; }
    }
}
