namespace ArmadaPack
{
    public struct CompareResult
    {
        public string Message { get; set; }
        public bool ProcessResult { get; set; }
        public string PercentResult { get; set; }

        public bool IsPass(double passRate = 0.05)
        {
            if (!ProcessResult) return false;

            var r = double.TryParse(Message, out var num);
            if (r)
            {
                PercentResult = num.ToString("P");
                return num < passRate;
            }

            return false;
        }
    }
}