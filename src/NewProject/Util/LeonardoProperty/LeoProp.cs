using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using NewProject.Util.PartList;

namespace NewProject.Util.LeonardoProperty;

public static class LeoProp
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

    private static string GetProjectDirectory(LeonardoType type, string program, string dir)
    {
        var hierarchy = GetLeoHierarchy(type);
        var segments = program.Split('\\', 3);
        var disk = segments[0];
        var archive = segments[1];
        var project = segments[2];

        return Path.Combine(disk, hierarchy, dir, archive, project);
    }

    public static string GetAdapterDirectory(LeonardoType type, string program)
    {
        return GetProjectDirectory(type, program, "ADAPTER");
    }

    public static string GetBoardDirectory(LeonardoType type, string program)
    {
        return GetProjectDirectory(type, program, "BOARD");
    }

    public static Side GetTopSide(LeonardoType type, string program)
    {
        var adapterDir = GetAdapterDirectory(type, program);
        var iniFile = Path.Combine(adapterDir, "ADAPTER.INI");
        var ini = new Ini(iniFile);

        var side = ini.ReadOne("UUT Position", "TopSide");

        return string.Equals(side, "Up", StringComparison.OrdinalIgnoreCase) ? Side.Top : Side.Bottom;
    }

    public static Dictionary<string, string> GetDeviceCode()
    {
        return File.ReadLines(
            Path.Combine(Directory.GetDirectories(@"C:\", "LEO*")
                .First(d => Regex.IsMatch(d, LeoDirRegexPattern)), "CADPACK\\BOM", DeviceTypeFile))
            .Where(l => l.Length > 0)
            .Select(s => s.Split(';'))
            .ToDictionary(s => s[0], s => s[1]);
    }

    public static Dictionary<string, string> GetBom(string boardDir)
    {
        var bomFile = Path.Combine(boardDir, "Bom.csv");

        return File.ReadLines(bomFile)
            .Where(l => l.Length > 0)
            .Select(s => s.Split(';', StringSplitOptions.TrimEntries))
            .ToDictionary(s => s[0], s => s[4]);
    }
}