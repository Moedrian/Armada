using System;

namespace ArmadaPack
{
    [Serializable]
    public class TestPart
    {
        public string DrawingReference { get; set; } = string.Empty;
        public double BarycenterX { get; set; }
        public double BarycenterY { get; set; }
        public int CoordinateX { get; set; }
        public int CoordinateY { get; set; }
        public int Edge { get; set; }
        public string GoldenSampleFilePath { get; set; }
    }
}
