using System;
using NewProject.Util.Dbf;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using NewProject.Util.LeonardoProperty;
using NewProject.Util.TpList;

namespace NewProject.Util.PartList;

public class Part
{
    public required string DrawingReference { get; init; }
    public required Side MountSide { get; init; }
    public required Barycenter Barycenter { get; init; } = new();
    public string DeviceType { get; private set; } = "Not Identified";
    public required bool Available { get; init; }
    public Coordinates Coordinates { get; set; } = new();

    [SetsRequiredMembers]
    public Part(DbfRecord dbfRecord)
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

            Available = true;
        }
        else
        {
            Available = false;
        }
    }

    public bool TryCalculateBarycenterCoordinates(TestPoint[] testPoints, out Coordinates co)
    {
        var dr = testPoints
            .Where(t => string.Equals(t.DrawingReference, DrawingReference, StringComparison.OrdinalIgnoreCase))
            .ToArray();

        if (dr.Length == 0)
        {
            co = new Coordinates();
            return false;
        }

        var x = (int)Math.Truncate(dr.Select(r => r.XPosition).Average());

        var xMin = dr.Select(r => r.XPosition).Min();
        var xMax = dr.Select(r => r.XPosition).Max();
        var xDiff = xMax - xMin;

        var y = (int)Math.Truncate(dr.Select(r => r.YPosition).Average());

        var yMin = dr.Select(r => r.YPosition).Min();
        var yMax = dr.Select(r => r.YPosition).Max();
        var yDiff = yMax - yMin;
        var diff = (int)Math.Truncate((new[] { xDiff, yDiff }).Max());

        co =  new Coordinates { X = x, Y = y, Edge = diff };

        return true;
    }

    public void SetDeviceType(string type)
    {
        DeviceType = type;
    }

    public static Part[] GetAllComponents(LeonardoType type, string program, out bool bomExists)
    {
        var boardDir = LeoProp.GetBoardDirectory(type, program);
        var dbfFile = Path.Combine(boardDir, "PARTLIST.DBF");
        var parser = new DbfParser(dbfFile);
        var records = parser.GetAllDbfRecords();
        var cl = new List<Part>();
        Dictionary<string, string> bom = [];

        bomExists = false;

        try
        {
            bom = LeoProp.GetBom(boardDir);
            bomExists = true;
        }
        catch (FileNotFoundException)
        {
            // using for marking bom file's non-existence
        }

        var availableComponents = records.Select(r => new Part(r)).Where(c => c.Available);

        if (!bomExists) return [.. availableComponents];

        foreach (var aC in availableComponents)
        {
            var dr = aC.DrawingReference;

            if (!bom.TryGetValue(dr, out var value)) continue;
            aC.SetDeviceType(value);
            cl.Add(aC);
        }

        return [.. cl];
    }

    public static Part[] GetTopSideComponents(LeonardoType type, string program, out bool bomExists)
    {
        var side = LeoProp.GetTopSide(type, program);
        return GetAllComponents(type, program, out bomExists).Where(c => c.MountSide == side).ToArray();
    }
}