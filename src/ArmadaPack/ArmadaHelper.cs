using System;
using System.Diagnostics;
using System.IO;
using System.Xml.Serialization;

namespace ArmadaPack
{
    public static class ArmadaHelper
    {
        public static XmlSerializer VisionSerializer = new XmlSerializer(typeof(Vision));

        public static int MillimetersToPixels(int len, int pixelRatio, int offsetPixels)
        {
            return (int)Math.Truncate((double)len / pixelRatio) + offsetPixels * 2;
        }

        public static void Serialize(string filename, Vision vision)
        {
            using (var sw = new StreamWriter(filename))
                VisionSerializer.Serialize(sw, vision);
        }

        public static Vision Deserialize(string filename)
        {
            using (var sr = new StreamReader(File.OpenRead(filename)))
            {
                var obj = VisionSerializer.Deserialize(sr) as Vision
                          ?? throw new Exception($"NULL content read in the file {filename}.");

                return obj;
            }
        }

        public static CompareResult CompareImages(string workingDirectory, string goldenBoardPic, string tookPic, int edge)
        {
            var cr = new CompareResult();

            try
            {
                var interpreter = Path.Combine(workingDirectory, "py_env", "pythonw.exe");
                var script = Path.Combine(workingDirectory, "py_env", "main.py");
                var psi = new ProcessStartInfo
                {
                    FileName = interpreter,
                    Arguments = $"\"{script}\" \"{goldenBoardPic}\" \"{tookPic}\" \"{edge}\"",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false
                };

                var p = Process.Start(psi) ?? throw new Exception($"Process {psi.FileName} is null.");
                p.WaitForExit();

                if (p.ExitCode != 0)
                {
                    cr.ProcessResult = false;
                    cr.Message = p.StandardError.ReadToEnd();
                }
                else
                {
                    cr.ProcessResult = true;
                    cr.Message = p.StandardOutput.ReadToEnd();
                }
            }
            catch (Exception e)
            {
                cr.ProcessResult = false;
                cr.Message = e.Message;
            }

            return cr;
        }
    }
}