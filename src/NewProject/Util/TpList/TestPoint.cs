using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using NewProject.Util.Dbf;
using NewProject.Util.LeonardoProperty;
using NewProject.Util.PartList;

namespace NewProject.Util.TpList;

public class TestPoint
{
    public required string Name { get; init; }
    public required string DrawingReference { get; init; }
    public required string Number { get; init; }
    public required double XPosition { get; init; }
    public required double YPosition { get; init; }
    public required bool Available { get; init; }

    [SetsRequiredMembers]
    public TestPoint(DbfRecord record)
    {
        const string tName = "TP_NAME";
        const string dName = "DRAWING_RE";
        const string xName = "X_POSITION";
        const string yName = "Y_POSITION";
        const string numName = "TP_NUM";

        var dict = record.FieldValues;
        Name = dict[tName];
        Number = dict[numName];
        DrawingReference = dict[dName];
        var xR = double.TryParse(dict[xName], out var x);
        var yR = double.TryParse(dict[yName], out var y);

        if (xR && yR)
        {
            XPosition = Math.Truncate(x * 1000);
            YPosition = Math.Truncate(y * 1000);
            Available = true;
        }
        else
        {
            Available = false;
        }
    }

    public static TestPoint[] GetAllTestPoints(LeonardoType type, string program)
    {
        var adapterDir = LeoProp.GetAdapterDirectory(type, program);
        var dbfFile = Path.Combine(adapterDir, "TPLIST.DBF");
        var parser = new DbfParser(dbfFile);
        var records = parser.GetAllDbfRecords();

        var testPoints = records.Select(r => new TestPoint(r)).Where(tp => tp.Available);

        return testPoints.ToArray();
    }

    public override string ToString()
    {
        return $"Name: {Name}, Number: {Number}, Drawing Reference: {DrawingReference}, X: {XPosition}, Y: {YPosition}";
    }
}