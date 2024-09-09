using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using NewProject.Util.Dbf;

namespace NewProject.Util.PartList;

public static class PartListGetter
{
    private const string LeoDirRegexPattern = @"C:\\LEO(OS2\.|YA)[0-9]{3}(P[A-Z])?";
    private const string DeviceTypeFile = "DevTypeSpeaCode.csv";

    private static string GetLeoHierarchy(LeonardoType type)
    {
        string hierarchy;
        if (type == LeonardoType.BedOfNails)
            hierarchy = "Leonardo\\TPRJ110";
        else if (type == LeonardoType.FlyingProbes)
            hierarchy = "LeonardoOS2\\FlyTprj.100";
        else
            throw new InvalidOperationException("Type not supported.");

        return hierarchy;
    }

    private static Side GetTopSide(LeonardoType type, string program)
    {
        var hierarchy = GetLeoHierarchy(type);
        var segments = program.Split('\\', 3);
        var disk = segments[0];
        var archive = segments[1];
        var project = segments[2];

        var iniFile = Path.Combine(disk, hierarchy, "ADAPTER", archive, project, "ADAPTER.INI");
        var ini = new Ini(iniFile);

        var side = ini.ReadOne("UUT Position", "TopSide");

        return string.Equals(side, "Up", StringComparison.OrdinalIgnoreCase) ? Side.Top : Side.Bottom;
    }

    private static string GetBoardDirectory(LeonardoType type, string program)
    {
        var hierarchy = GetLeoHierarchy(type);
        var segments = program.Split('\\', 3);
        var disk = segments[0];
        var archive = segments[1];
        var project = segments[2];

        return Path.Combine(disk, hierarchy, "BOARD", archive, project);
    }

    public static Dictionary<string, string> GetDeviceCode()
    {
        return File.ReadLines(
            Path.Combine(Directory .GetDirectories(@"C:\", "LEO*")
                .First(d => Regex.IsMatch(d, LeoDirRegexPattern)), "CADPACK\\BOM", DeviceTypeFile))
            .Where(l => l.Length > 0)
            .Select(s => s.Split(';'))
            .ToDictionary(s => s[0], s => s[1]);
    }

    private static Dictionary<string, string> GetBom(string boardDir)
    {
        var bomFile = Path.Combine(boardDir, "Bom.csv");

        return File.ReadLines(bomFile)
            .Where(l => l.Length > 0)
            .Select(s => s.Split(';', StringSplitOptions.TrimEntries))
            .ToDictionary(s => s[0], s => s[4]);
    }

    public static Component[] GetAllComponents(LeonardoType type, string program, out bool bomExists)
    {
        var boardDir = GetBoardDirectory(type, program);
        var dbfFile = Path.Combine(boardDir, "PARTLIST.DBF");
        var parser = new DbfParser(dbfFile);
        var records = parser.GetAllDbfRecords();
        var cl = new List<Component>();
        Dictionary<string, string> bom = [];

        bomExists = false;

        try
            {
            bom = GetBom(boardDir);
            bomExists = true;
        }
        catch (FileNotFoundException)
        {
            // hmmm....
        }

        var availableComponents = records.Select(r => new Component(r)).Where(c => c.Available);

        if (!bomExists) return [..availableComponents];

        foreach (var aC in availableComponents)
        {
            var dr = aC.DrawingReference;

            if (!bom.TryGetValue(dr, out var value)) continue;
            aC.SetDeviceType(value);
            cl.Add(aC);
        }

        return [..cl];
    }

    public static Component[] GetTopSideComponents(LeonardoType type, string program, out bool bomExists)
    {
        var side = GetTopSide(type, program);
        return GetAllComponents(type, program, out bomExists).Where(c => c.MountSide == side).ToArray();
    }
}