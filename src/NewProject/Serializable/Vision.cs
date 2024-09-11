using System;
using System.Xml.Serialization;
using NewProject.Util.PartList;

namespace NewProject.Serializable;

[Serializable]
public record Vision
{
    [XmlArray("TestParts")]
    public TestPart[] TestParts { get; set; } = [];
}