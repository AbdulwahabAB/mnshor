using System.Diagnostics;
using System.Runtime.InteropServices;
using System.ServiceProcess;

namespace Mnshor.Launcher;

internal static class Program
{
    private const string ServiceName = "MnshorService";
    private const string ServiceProbeUrl = "http://127.0.0.1:5157/";
    private const string BrowserLaunchUrl = "http://localhost:5157";

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int MessageBox(IntPtr hWnd, string text, string caption, uint type);

    [STAThread]
    private static async Task Main()
    {
        using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(1.5) };

        if (!await IsEndpointReadyAsync(http))
        {
            var started = TryStartWindowsService() || TryStartLocalProcess();
            if (started)
            {
                await WaitForEndpointAsync(http, TimeSpan.FromSeconds(10));
            }
        }

        if (!await IsEndpointReadyAsync(http))
        {
            MessageBox(
                IntPtr.Zero,
                $"Mnshor could not connect to the local service ({BrowserLaunchUrl}).\n\nPlease check that \"{ServiceName}\" is running in Windows Services.",
                "Mnshor",
                0x10 /* MB_ICONERROR */);
            return;
        }

        Process.Start(new ProcessStartInfo(BrowserLaunchUrl) { UseShellExecute = true });
    }

    private static async Task<bool> IsEndpointReadyAsync(HttpClient http)
    {
        try
        {
            using var response = await http.GetAsync(ServiceProbeUrl);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static async Task WaitForEndpointAsync(HttpClient http, TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline)
        {
            if (await IsEndpointReadyAsync(http))
            {
                return;
            }

            await Task.Delay(300);
        }
    }

    private static bool TryStartWindowsService()
    {
        try
        {
            using var sc = new ServiceController(ServiceName);
            if (sc.Status == ServiceControllerStatus.Running)
            {
                return true;
            }

            if (sc.Status == ServiceControllerStatus.Stopped)
            {
                sc.Start();
            }

            sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(10));
            return sc.Status == ServiceControllerStatus.Running;
        }
        catch
        {
            return false;
        }
    }

    private static bool TryStartLocalProcess()
    {
        var exePath = Path.Combine(AppContext.BaseDirectory, "Mnshor.exe");
        if (!File.Exists(exePath))
        {
            return false;
        }

        try
        {
            Process.Start(new ProcessStartInfo(exePath)
            {
                WorkingDirectory = AppContext.BaseDirectory,
                CreateNoWindow = true,
                UseShellExecute = false
            });
            return true;
        }
        catch
        {
            return false;
        }
    }
}
