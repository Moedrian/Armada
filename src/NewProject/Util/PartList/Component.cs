using NewProject.Util.Dbf;

namespace NewProject.Util.PartList;

public class Component
{
    public string DrawingReference { get; init; }
    public Side MountSide { get; init; }
    public Barycenter Barycenter { get; init; } = new();
    public string DeviceType { get; private set; } = "Not Identified";
    public bool Available { get; init; } = true;

    public Component(DbfRecord dbfRecord)
    {
        const string dName = "DRAWING_RE";
        const string xName = "X_BARYCENT";
        const string yName = "Y_BARYCENT";
        const string sName = "MOUNT_SIDE";

        var dict = dbfRecord.FieldValues;
        DrawingReference = dict[dName];
        MountSide = dict[sName] == "B" ? Side.Bottom : Side.Top;
        var xR = double.TryParse(dict[xName], out var x);
        var yR = double.TryParse(dict[yName], out var y);

        if (xR && yR)
        {
            Barycenter = new Barycenter
            {
                X = x,
                Y = y,
            };
        }
        else
        {
            Available = false;
        }
    }

    public void SetDeviceType(string type)
    {
        DeviceType = type;
    }
}