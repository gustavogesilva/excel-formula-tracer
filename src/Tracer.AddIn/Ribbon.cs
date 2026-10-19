using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;
using Tracer.AddIn.Excel;
using Tracer.AddIn.UI;

namespace Tracer.AddIn;

/// <summary>A small "Tracer" ribbon tab; everything is also reachable by keyboard.</summary>
[ComVisible(true)]
public sealed class Ribbon : ExcelRibbon
{
    public override string GetCustomUI(string ribbonId) => @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='ftTab' label='Tracer'>
        <group id='ftTrace' label='Trace'>
          <button id='ftPrec' label='Precedents' size='large' imageMso='TracePrecedents' onAction='OnPrecedents' screentip='Trace precedents (Ctrl+Q)'/>
          <button id='ftDep' label='Dependents' size='large' imageMso='TraceDependents' onAction='OnDependents' screentip='Trace dependents (Ctrl+Shift+Q)'/>
          <button id='ftError' label='Trace error' size='large' imageMso='TraceError' onAction='OnError' screentip='Find the cell(s) where the error in the active cell starts (E in the trace window)'/>
          <button id='ftBack' label='Jump back' imageMso='GoLeftToRight' onAction='OnBack' screentip='Return to where you jumped from (Ctrl+Backspace)'/>
        </group>
        <group id='ftOptions' label='Options'>
          <dropDown id='ftKeys' label='Shortcuts' getSelectedItemIndex='GetKeysIndex' onAction='OnKeys'>
            <item id='kArixcel' label='Ctrl+Q / Ctrl+Shift+Q'/>
            <item id='kBrackets' label='Ctrl+[ / Ctrl+]'/>
            <item id='kBoth' label='Both'/>
          </dropDown>
          <checkBox id='ftPrebuild' label='Index at idle' getPressed='GetPrebuild' onAction='OnPrebuild' screentip='Build the dependents index in the background after a workbook opens'/>
          <button id='ftRebuild' label='Rebuild index' imageMso='Refresh' onAction='OnRebuild'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

    public void OnPrecedents(IRibbonControl control) => Commands.TracePrecedents();
    public void OnDependents(IRibbonControl control) => Commands.TraceDependents();
    public void OnError(IRibbonControl control) => Commands.TraceError();
    public void OnBack(IRibbonControl control) => Commands.JumpBack();
    public void OnRebuild(IRibbonControl control) => Commands.RebuildIndex();

    public int GetKeysIndex(IRibbonControl control) => (int)Settings.Current.Keys;

    public void OnKeys(IRibbonControl control, string selectedId, int selectedIndex)
    {
        Settings.Current.Keys = (KeyPreset)selectedIndex;
        Settings.Current.Save();
        Commands.Run(() => Hotkeys.Apply(Settings.Current.Keys));
    }

    public bool GetPrebuild(IRibbonControl control) => Settings.Current.Prebuild;

    public void OnPrebuild(IRibbonControl control, bool pressed)
    {
        Settings.Current.Prebuild = pressed;
        Settings.Current.Save();
        if (pressed) EventSync.SchedulePrebuild();
    }
}
