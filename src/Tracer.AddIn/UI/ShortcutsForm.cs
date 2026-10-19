using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Media;
using System.Windows.Forms;

namespace Tracer.AddIn.UI;

/// <summary>
/// Tracer › Shortcuts…: assign a primary and an alternate key to each tracing action. Shown modally, so WinForms'
/// own key pre-processing runs and each box can capture the combination the user presses.
/// </summary>
internal sealed class ShortcutsForm : Form
{
    private readonly Dictionary<(TraceAction Action, int Slot), ShortcutBox> _boxes = new();
    private readonly Label _hint;
    private const string DefaultHint =
        "Click a box and press the keys. Shortcuts need Ctrl or Alt (F1–F12 may be used alone). "
        + "While the add-in is loaded they replace whatever Excel does with those keys.";

    public ShortcutsForm(ShortcutMap current)
    {
        Result = current.Clone();
        Text = "Formula Tracer – Shortcuts";
        Font = new Font("Segoe UI", 9f);
        AutoScaleMode = AutoScaleMode.Font;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new Size(600, 312);

        Controls.Add(new Label { Text = "Shortcut", AutoSize = true, Location = new Point(200, 14), Font = new Font(Font, FontStyle.Bold) });
        Controls.Add(new Label { Text = "Alternate", AutoSize = true, Location = new Point(392, 14), Font = new Font(Font, FontStyle.Bold) });

        int y = 40;
        foreach (var action in ShortcutMap.Actions)
        {
            Controls.Add(new Label { Text = ShortcutMap.Label(action), AutoSize = true, Location = new Point(16, y + 4) });
            for (int slot = 0; slot < 2; slot++)
            {
                int x = slot == 0 ? 200 : 392;
                var box = new ShortcutBox { Location = new Point(x, y), Width = 150, Value = current[action, slot] };
                var key = (action, slot);
                box.Changed += () => Result[key.action, key.slot] = box.Value;
                box.Rejected += reason =>
                {
                    SystemSounds.Beep.Play();
                    _hint!.Text = reason;
                };
                var clear = new Button { Text = "×", Location = new Point(x + 152, y - 1), Size = new Size(28, 25), TabStop = false };
                clear.Click += (_, _) =>
                {
                    box.Value = Shortcut.None;
                    Result[key.action, key.slot] = Shortcut.None;
                };
                _boxes[key] = box;
                Controls.Add(box);
                Controls.Add(clear);
            }
            y += 36;
        }

        _hint = new Label { Text = DefaultHint, AutoSize = false, Location = new Point(16, y + 6), Size = new Size(568, 40), ForeColor = SystemColors.GrayText };
        Controls.Add(_hint);

        var arixcel = new Button { Text = "Arixcel keys", Location = new Point(16, 266), Size = new Size(110, 30) };
        var brackets = new Button { Text = "Excel brackets", Location = new Point(132, 266), Size = new Size(110, 30) };
        var ok = new Button { Text = "OK", Location = new Point(404, 266), Size = new Size(88, 30) };
        var cancel = new Button { Text = "Cancel", Location = new Point(498, 266), Size = new Size(88, 30), DialogResult = DialogResult.Cancel };
        arixcel.Click += (_, _) => ApplyPreset(ShortcutMap.Arixcel());
        brackets.Click += (_, _) => ApplyPreset(ShortcutMap.Brackets());
        ok.Click += (_, _) => Confirm();
        Controls.AddRange(new Control[] { arixcel, brackets, ok, cancel });
        AcceptButton = ok;
        CancelButton = cancel;
    }

    /// <summary>The edited shortcuts (valid when the dialog returns OK).</summary>
    public ShortcutMap Result { get; private set; }

    private void ApplyPreset(ShortcutMap map)
    {
        Result = map.Clone();
        foreach (var kv in _boxes) kv.Value.Value = Result[kv.Key.Action, kv.Key.Slot];
        _hint.Text = DefaultHint;
    }

    private void Confirm()
    {
        var problems = Result.Problems();
        if (problems.Count > 0)
        {
            MessageBox.Show(this, string.Join("\n", problems), Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        var overrides = Result.All()
            .Where(x => x.Key.ExcelConflict() != null)
            .Select(x => $"{x.Key} (Excel's {x.Key.ExcelConflict()}) for {ShortcutMap.Label(x.Action)}")
            .ToList();
        if (overrides.Count > 0
            && MessageBox.Show(this, "These replace common Excel shortcuts while the add-in is loaded:\n\n" + string.Join("\n", overrides) + "\n\nUse them anyway?",
                Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            return;
        DialogResult = DialogResult.OK;
        Close();
    }

    /// <summary>A read-only box that shows a shortcut and replaces it with the next key combination pressed.</summary>
    private sealed class ShortcutBox : TextBox
    {
        private Shortcut _value;

        public ShortcutBox()
        {
            ReadOnly = true;
            ShortcutsEnabled = false;
            BackColor = SystemColors.Window;
        }

        public event Action? Changed;
        public event Action<string>? Rejected;

        public Shortcut Value
        {
            get => _value;
            set
            {
                _value = value;
                Text = value.IsEmpty ? "(none)" : value.ToString();
            }
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            // Leave dialog navigation alone: Tab moves on, Enter confirms, Esc cancels.
            if (keyData == Keys.Tab || keyData == (Keys.Shift | Keys.Tab) || keyData == Keys.Enter || keyData == Keys.Escape)
                return base.ProcessCmdKey(ref msg, keyData);

            var code = keyData & Keys.KeyCode;
            if (code == Keys.ControlKey || code == Keys.ShiftKey || code == Keys.Menu || code == Keys.LWin || code == Keys.RWin)
                return true; // a modifier on its own: wait for the actual key

            var shortcut = Shortcut.FromKeyData((int)keyData);
            if (shortcut.IsEmpty)
            {
                Rejected?.Invoke($"{code} can't be used in a shortcut.");
                return true;
            }
            if (shortcut.Problem() is { } problem)
            {
                Rejected?.Invoke(problem);
                return true;
            }
            Value = shortcut;
            Changed?.Invoke();
            return true;
        }
    }
}

/// <summary>Lets a WinForms dialog be owned by Excel's (native) main window.</summary>
internal sealed class NativeOwnerWindow : IWin32Window
{
    public NativeOwnerWindow(IntPtr handle) => Handle = handle;

    public IntPtr Handle { get; }
}
