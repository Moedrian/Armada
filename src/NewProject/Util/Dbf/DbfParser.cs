using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NewProject.Util.Dbf;

public class DbfParser
{
    private readonly byte[] _rawBytes;
    private readonly int _headerLength;

    public readonly int RecordCount;
    public int NumberOfFields { get; }
    public string[] Fields { get; }
    public int[] FieldLengths { get; }

    private readonly List<List<byte>> _recordBytesLists = new();

    private const int HeaderLengthWithoutFieldDescriptor = 33;
    private const int FieldDescriptorLength = 32;
    private const int FieldDescriptorStartIdx = 32;

    public DbfParser(string filename)
    {
        _rawBytes = File.ReadAllBytes(filename);

        RecordCount = GetLittleEndianInteger(_rawBytes[4..8]);
        _headerLength = GetLittleEndianInteger(_rawBytes[8..10]);
        NumberOfFields = (_headerLength - HeaderLengthWithoutFieldDescriptor) / FieldDescriptorLength;

        var fn = new List<string>();
        var fl = new List<int>();
        for (var i = 0; i < NumberOfFields; i++)
        {
            var fieldNameStart = FieldDescriptorStartIdx + i * FieldDescriptorLength;
            var fieldNameEnd = fieldNameStart + 10;
            var fieldNameBytes = _rawBytes[fieldNameStart..fieldNameEnd].Where(x => x > 0).ToArray();
            fn.Add(Encoding.ASCII.GetString(fieldNameBytes));
            var lengthIdx = fieldNameStart + 16;
            fl.Add(_rawBytes[lengthIdx]);
        }

        Fields = fn.ToArray();
        FieldLengths = fl.ToArray();
    }

    private static int GetLittleEndianInteger(byte[] littleEndianByteArray)
    {
        var reversed = littleEndianByteArray.Reverse();
        var hexString = string.Join("", reversed.Select(x => x.ToString("X2")));
        return Convert.ToInt32(hexString, 16);
    }

    public DbfRecord[] GetAllDbfRecords()
    {
        var recordLength = FieldLengths.Sum() + 1;    // start byte marking deletion

        var l = new List<byte>();
        var counter = 0;
        for (var i = _headerLength; i < _rawBytes.Length; i++)
        {
            l.Add(_rawBytes[i]);
            counter++;
            if (counter == recordLength)
            {
                var cloneList = new List<byte>();
                cloneList.AddRange(l);
                _recordBytesLists.Add(cloneList);
                l.Clear();
                counter = 0;
            }
        }

        var dbfRecords = new List<DbfRecord>();
        foreach (var bs in _recordBytesLists)
            dbfRecords.Add(new DbfRecord(FieldLengths, Fields, bs.ToArray()));

        return dbfRecords.ToArray();
    }
}