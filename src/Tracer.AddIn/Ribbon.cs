using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Tracer.AddIn.Excel;
using Tracer.AddIn.UI;

namespace Tracer.AddIn;

/// <summary>A small "Tracer" ribbon tab; everything is also reachable by keyboard.</summary>
[ComVisible(true)]
public sealed class Ribbon : ExcelRibbon
{
    private IRibbonUI? _ui;

    public override string GetCustomUI(string ribbonId) => @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnLoad'>
  <ribbon>
    <tabs>
      <tab id='ftTab' label='Tracer'>
        <group id='ftTrace' label='Trace'>
          <button id='ftPrec' label='Precedents' size='large' imageMso='TracePrecedents' onAction='OnPrecedents' getScreentip='GetTip'/>
          <button id='ftDep' label='Dependents' size='large' imageMso='TraceDependents' onAction='OnDependents' getScreentip='GetTip'/>
          <button id='ftError' label='Trace error' size='large' imageMso='TraceError' onAction='OnError' getScreentip='GetTip'/>
          <button id='ftBack' label='Jump back' imageMso='GoLeftToRight' onAction='OnBack' getScreentip='GetTip'/>
        </group>
        <group id='ftOptions' label='Options'>
          <button id='ftShortcuts' label='Shortcuts…' onAction='OnShortcuts' screentip='Choose the keyboard shortcuts for tracing'/>
          <checkBox id='ftCrosshair' label='Highlight row &amp; column' getPressed='GetCrosshair' onAction='OnCrosshair' screentip='While tracing, shade the row and column of the cell you move to'/>
          <checkBox id='ftPrebuild' label='Index at idle' getPressed='GetPrebuild' onAction='OnPrebuild' screentip='Build the dependents index in the background after a workbook opens'/>
          <button id='ftRebuild' label='Rebuild index' imageMso='Refresh' onAction='OnRebuild'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

    public void OnLoad(IRibbonUI ui) => _ui = ui;

    public void OnPrecedents(IRibbonControl control) => Commands.TracePrecedents();
    public void OnDependents(IRibbonControl control) => Commands.TraceDependents();
    public void OnError(IRibbonControl control) => Commands.TraceError();
    public void OnBack(IRibbonControl control) => Commands.JumpBack();
    public void OnRebuild(IRibbonControl control) => Commands.RebuildIndex();

    /// <summary>Tooltips show the shortcuts currently assigned.</summary>
    public string GetTip(IRibbonControl control)
    {
        var action = control.Id switch
        {
            "ftPrec" => TraceAction.TracePrecedents,
            "ftDep" => TraceAction.TraceDependents,
            "ftError" => TraceAction.TraceError,
            _ => TraceAction.JumpBack,
        };
        string keys = Settings.Current.Shortcuts.Describe(action);
        return ShortcutMap.Label(action) + (keys.Length > 0 ? $" ({keys})" : " (no shortcut: set one in Shortcuts…)");
    }

    public void OnShortcuts(IRibbonControl control) => Commands.Run("changing shortcuts", () =>
    {
        using var form = new ShortcutsForm(Settings.Current.Shortcuts);
        if (form.ShowDialog(new NativeOwnerWindow(ExcelDnaUtil.WindowHandle)) != DialogResult.OK) return;
        Settings.Current.Shortcuts = form.Result;
        Settings.Current.Save();
        var refused = Hotkeys.Apply(form.Result);
        _ui?.Invalidate();
        if (refused.Count > 0)
            MessageBox.Show("Excel did not accept these shortcuts:\n\n" + string.Join("\n", refused), "Formula Tracer",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
    });

    public bool GetCrosshair(IRibbonControl control) => Settings.Current.HighlightCrosshair;

    public void OnCrosshair(IRibbonControl control, bool pressed)
    {
        Settings.Current.HighlightCrosshair = pressed;
        Settings.Current.Save();
        if (!pressed) TraceController.Instance.Window.HideCrosshair();
    }

    public bool GetPrebuild(IRibbonControl control) => Settings.Current.Prebuild;

    public void OnPrebuild(IRibbonControl control, bool pressed)
    {
        Settings.Current.Prebuild = pressed;
        Settings.Current.Save();
        if (pressed) EventSync.SchedulePrebuild();
    }
}
