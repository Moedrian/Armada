using System;
using System.Text.Json.Serialization;
using NewProject.Util.PartList;

namespace NewProject.Serializable;

[Serializable]
public class CheckComponent
{
    [JsonIgnore]
    public bool IsChecked { get; set; }

    public string ComponentName { get; set; } = string.Empty;
    public Barycenter Barycenter { get; set; } = new();

    [JsonIgnore]
    public Component? Component { get; set; }

    public CheckComponent()
    { }

    public CheckComponent(Component component)
    {
        Component = component;
        ComponentName = component.DrawingReference;
        Barycenter = component.Barycenter;
    }
}