using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

internal static class Launcher {
    [STAThread] static void Main() {
        try {
            string root = AppDomain.CurrentDomain.BaseDirectory;
            string ps = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), @"PowerShell\7\pwsh.exe");
            if (!File.Exists(ps)) ps = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), @"System32\WindowsPowerShell\v1.0\powershell.exe");
            var start = new ProcessStartInfo(ps, "-NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File \"" + Path.Combine(root, "Launch-PolishingSentinel.ps1") + "\"");
            start.UseShellExecute = false;
            start.CreateNoWindow = true;
            start.WorkingDirectory = root;
            Process.Start(start);
        } catch (Exception ex) { MessageBox.Show(ex.Message, "Polishing Sentinel"); }
    }
}
