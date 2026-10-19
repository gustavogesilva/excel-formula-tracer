using System;
using System.Linq;
using System.Windows.Forms;

namespace Tracer.Setup;

/// <summary>
/// <c>FormulaTracerSetup.exe</c>: no arguments shows the setup dialog. <c>/install</c> or <c>/uninstall</c> act directly,
/// and with <c>/quiet</c> no UI is shown. Exit codes for /quiet: 0 done, 1 failed, 2 Excel is running, 3 Excel not found.
/// </summary>
internal static class Program
{
    public const int Ok = 0, Failed = 1, ExcelRunning = 2, ExcelNotFound = 3;

    [STAThread]
    private static int Main(string[] args)
    {
        bool Has(string flag) => args.Any(a => string.Equals(a.TrimStart('/', '-'), flag, StringComparison.OrdinalIgnoreCase));
        bool quiet = Has("quiet") || Has("q");
        bool uninstall = Has("uninstall");
        bool install = Has("install");

        int code;
        try
        {
            if (quiet)
            {
                code = uninstall ? QuietUninstall() : QuietInstall();
            }
            else
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                if (uninstall) code = SetupForm.ConfirmAndUninstall(owner: null) ? Ok : Failed;
                else
                {
                    Application.Run(new SetupForm(autoInstall: install));
                    code = Ok;
                }
            }
        }
        finally
        {
            if (Installer.PendingFolderDelete) Installer.DeleteFolderAfterExit();
        }
        return code;
    }

    private static int QuietInstall()
    {
        if (ExcelLocator.IsRunning()) return ExcelRunning;
        var excel = ExcelLocator.Find();
        if (excel == null) return ExcelNotFound;
        try
        {
            Installer.Install(excel);
            return Ok;
        }
        catch (Exception)
        {
            return Failed;
        }
    }

    private static int QuietUninstall()
    {
        if (ExcelLocator.IsRunning()) return ExcelRunning;
        try
        {
            Installer.Uninstall();
            return Ok;
        }
        catch (Exception)
        {
            return Failed;
        }
    }
}
