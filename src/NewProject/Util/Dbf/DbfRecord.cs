using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NewProject.Util.Dbf;

public class DbfRecord
{
    public bool Deleted { get; }
    public readonly Dictionary<string, string> FieldValues = new();

    public DbfRecord(int[] fieldLengths, string[] fieldNames, byte[] bytes)
    {
        Deleted = bytes[0] == 0x2A;
        var startIdx = 1;
        for (var i = 0; i < fieldLengths.Length; i++)
        {
            var endIdx = startIdx + fieldLengths[i];
            var fieldValue = bytes[startIdx..endIdx].Where(x => x > 0x20).ToArray();
            FieldValues.Add(fieldNames[i], Encoding.ASCII.GetString(fieldValue));
            startIdx += fieldLengths[i];
        }
    }
}