using System.Diagnostics;
using System.Text.Json;
using ArmadaPack;

namespace PyRunner;

internal static class Program
{
    private static readonly string Dir = AppDomain.CurrentDomain.BaseDirectory;
    private static readonly string SettingsFile = Path.Combine(Dir, "py_runner_settings.json");
    private static readonly JsonSerializerOptions Jso = new() { WriteIndented = true };

    private static int Main(string[] args)
    {
        if (args.Length != 2)
        {
            Console.Error.Write("Args length not correct.");
            return -1;
        }

        try
        {
            if (!File.Exists(SettingsFile))
            {
                var initialSettings = new PyRunnerSettings
                {
                    Interpreter = Path.Combine(Dir,"py_env", "python.exe"),
                    Script = Path.Combine(Dir, "py_scripts", "main.py")
                };
                var json = JsonSerializer.Serialize(initialSettings, Jso);
                File.WriteAllText(SettingsFile, json);
            }

            var jsonRead = File.ReadAllText(SettingsFile);
            var settings = JsonSerializer.Deserialize<PyRunnerSettings>(jsonRead)
                           ?? throw new Exception($"null content in the {SettingsFile}.");

            var pic1 = "\"" + args[0] + "\"";
            var pic2 = "\"" + args[1] + "\"";

            var psi = new ProcessStartInfo(settings.Interpreter, [settings.Script, pic1, pic2])
            {
                RedirectStandardError = true,
                RedirectStandardOutput = true,
                UseShellExecute = false
            };

            var p = Process.Start(psi) ?? throw new Exception("Unable to start python process.");
            p.WaitForExit();

            if (p.ExitCode != 0)
            {
                var stdErr = p.StandardError.ReadToEnd();
                throw new Exception(stdErr);
            }

            Console.Out.Write(p.StandardOutput.ReadToEnd());

            return 0;
        }
        catch (Exception e)
        {
            Console.Error.Write(e.Message);
            return -1;
        }
    }
}
