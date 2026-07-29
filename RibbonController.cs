using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace ExcelSheetManager
{
    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        public override string GetCustomUI(string ribbonId)
        {
            return @"
<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
  <ribbon>
    <tabs>
      <tab id='tabExcelSheetManager' label='Sheet Manager'>
        <group id='groupManager' label='Navigation'>
          <button id='btnToggleTaskPane' 
                  label='Toggle Taskpane' 
                  size='large' 
                  imageMso='SheetManager' 
                  onAction='OnToggleTaskPane' />
          <button id='btnRefreshTaskPane' 
                  label='Refresh Lists' 
                  size='large' 
                  imageMso='Refresh' 
                  onAction='OnRefreshTaskPane' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void OnToggleTaskPane(IRibbonControl control)
        {
            AddIn.ToggleTaskPane();
        }

        public void OnRefreshTaskPane(IRibbonControl control)
        {
            AddIn.RefreshTaskPane();
        }
    }
}
