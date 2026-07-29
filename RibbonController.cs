using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
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
                  imageMso='GroupWorksheet' 
                  onAction='OnToggleTaskPane' />
          <button id='btnRefreshTaskPane' 
                  label='Refresh Lists' 
                  size='large' 
                  imageMso='RefreshAll' 
                  onAction='OnRefreshTaskPane' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void OnToggleTaskPane(IRibbonControl control)
        {
            try
            {
                AddIn.ToggleTaskPane();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Toggle Taskpane Error: {ex.Message}\n\n{ex.StackTrace}", "Excel Sheet Manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnRefreshTaskPane(IRibbonControl control)
        {
            try
            {
                AddIn.RefreshTaskPane();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Refresh Error: {ex.Message}\n\n{ex.StackTrace}", "Excel Sheet Manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
