using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace Tracer.Setup;

/// <summary>The setup dialog: shows the detected Excel and install state, and installs, updates or uninstalls.</summary>
internal sealed class SetupForm : Form
{
    private readonly Label _excel;
    private readonly Label _state;
    private readonly Button _install;
    private readonly Button _uninstall;
    private readonly bool _autoInstall;
    private ExcelInstall? _found;

    public SetupForm(bool autoInstall)
    {
        _autoInstall = autoInstall;
        Text = "Formula Tracer Setup";
        Font = new Font("Segoe UI", 9f);
        AutoScaleMode = AutoScaleMode.Font;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new Size(540, 250);

        var title = new Label
        {
            Text = Installer.DisplayName,
            Font = new Font("Segoe UI", 14f, FontStyle.Bold),
            AutoSize = true,
            Location = new Point(20, 16),
        };
        var version = new Label { Text = "Version " + Installer.Version, AutoSize = true, ForeColor = SystemColors.GrayText, Location = new Point(23, 52) };
        _excel = new Label { AutoSize = false, Location = new Point(23, 84), Size = new Size(500, 36) };
        _state = new Label { AutoSize = false, Location = new Point(23, 124), Size = new Size(500, 36) };
        var note = new Label
        {
            Text = "Installs for your Windows user only (no administrator rights). Excel must be closed. "
                 + "After installing, the Tracer tab appears on Excel's ribbon.",
            AutoSize = false,
            ForeColor = SystemColors.GrayText,
            Location = new Point(23, 164),
            Size = new Size(500, 36),
        };

        _install = new Button { Text = "Install", Size = new Size(100, 30), Location = new Point(204, 208) };
        _uninstall = new Button { Text = "Uninstall", Size = new Size(100, 30), Location = new Point(314, 208) };
        var close = new Button { Text = "Close", Size = new Size(100, 30), Location = new Point(424, 208), DialogResult = DialogResult.Cancel };
        _install.Click += (_, _) => DoInstall();
        _uninstall.Click += (_, _) =>
        {
            if (ConfirmAndUninstall(this)) RefreshState();
        };
        AcceptButton = _install;
        CancelButton = close;

        Controls.AddRange(new Control[] { title, version, _excel, _state, note, _install, _uninstall, close });
        RefreshState();
    }

    protected override void OnShown(EventArgs e)
    {
        base.OnShown(e);
        if (_autoInstall && _install.Enabled) DoInstall();
    }

    private void RefreshState()
    {
        _found = ExcelLocator.Find();
        _excel.Text = _found == null
            ? "Microsoft Excel was not found on this computer."
            : $"Found {_found}\n{_found.ExePath}";
        _excel.ForeColor = _found == null ? Color.Firebrick : SystemColors.ControlText;

        string? installed = Installer.InstalledVersion();
        _state.Text = installed == null
            ? "Formula Tracer is not installed."
            : $"Formula Tracer {installed} is installed in {Installer.InstallDir}.";
        _install.Text = installed == null ? "Install" : "Update";
        _install.Enabled = _found != null;
        _uninstall.Enabled = installed != null;
    }

    private void DoInstall()
    {
        if (_found == null || !WaitForExcelToClose(this, "install")) return;
        try
        {
            Installer.Install(_found);
            MessageBox.Show(this,
                "Formula Tracer is installed.\n\nStart Excel: the Tracer tab is on the ribbon.\n"
                + "Ctrl+Q traces precedents, Ctrl+Shift+Q traces dependents.",
                Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is InvalidOperationException || ex is System.Security.SecurityException)
        {
            MessageBox.Show(this, "Installation failed:\n\n" + ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        RefreshState();
    }

    /// <summary>Asks for confirmation and Excel to be closed, then uninstalls. Also used for "Uninstall" in Apps &amp; features.</summary>
    public static bool ConfirmAndUninstall(IWin32Window? owner)
    {
        const string caption = "Formula Tracer Setup";
        if (MessageBox.Show(owner, "Remove Formula Tracer from Excel?", caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            return false;
        if (!WaitForExcelToClose(owner, "uninstall")) return false;
        try
        {
            Installer.Uninstall();
            MessageBox.Show(owner, "Formula Tracer has been removed.", caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            return true;
        }
        catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is System.Security.SecurityException)
        {
            MessageBox.Show(owner, "Uninstall failed:\n\n" + ex.Message, caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }
    }

    /// <summary>Excel rewrites its add-in list when it exits, so changes made while it runs would be lost.</summary>
    private static bool WaitForExcelToClose(IWin32Window? owner, string action)
    {
        while (ExcelLocator.IsRunning())
        {
            var answer = MessageBox.Show(owner,
                $"Please close Excel to {action} Formula Tracer (save your work first), then click Retry.",
                "Formula Tracer Setup", MessageBoxButtons.RetryCancel, MessageBoxIcon.Warning);
            if (answer != DialogResult.Retry) return false;
        }
        return true;
    }
}
