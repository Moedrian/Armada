using System;

namespace NewProject.Serializable;

[Serializable]
public record TestPart
{
    public string DrawingReference { get; set; } = string.Empty;
    public double X { get; set; }
    public double Y { get; set; }
}