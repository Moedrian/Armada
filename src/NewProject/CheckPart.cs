using System;
using NewProject.Util.PartList;

namespace NewProject;

[Serializable]
public class CheckPart(Part part)
{
    public bool IsChecked { get; set; }
    public Part? Part { get; set; } = part;
    public string PartName { get; set; } = part.DrawingReference;
    public Barycenter Barycenter { get; set; } = part.Barycenter;
    public Coordinates Coordinates { get; set; } = part.Coordinates;
}