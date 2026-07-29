using System;
using System.Drawing;
using System.Drawing.Drawing2D;
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
                  getImage='GetRibbonImage' 
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

        public Bitmap? GetRibbonImage(IRibbonControl control)
        {
            if (control.Id == "btnToggleTaskPane")
            {
                return CreateTaskPaneIcon();
            }
            return null;
        }

        private Bitmap CreateTaskPaneIcon()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;

                // Main Outer Border / Background
                using (var bgBrush = new SolidBrush(Color.FromArgb(15, 23, 42))) // Dark slate background
                {
                    g.FillRectangle(bgBrush, 2, 2, 28, 28);
                }

                using (var borderPen = new Pen(Color.FromArgb(56, 189, 248), 2)) // Sky blue border
                {
                    g.DrawRectangle(borderPen, 2, 2, 28, 28);
                }

                // Left region lines representing Excel Sheet Rows
                using (var linePen = new Pen(Color.FromArgb(148, 163, 184), 1.5f))
                {
                    g.DrawLine(linePen, 6, 9, 17, 9);
                    g.DrawLine(linePen, 6, 16, 17, 16);
                    g.DrawLine(linePen, 6, 23, 17, 23);
                }

                // Divider line
                using (var divPen = new Pen(Color.FromArgb(51, 65, 85), 1))
                {
                    g.DrawLine(divPen, 19, 4, 19, 28);
                }

                // Right Taskpane accent panel bar
                using (var paneBrush = new SolidBrush(Color.FromArgb(2, 132, 199))) // Vivid accent blue
                {
                    g.FillRectangle(paneBrush, 21, 5, 7, 22);
                }

                // Taskpane sheet tabs representation (small colored dots)
                using (var dot1 = new SolidBrush(Color.FromArgb(56, 189, 248)))
                using (var dot2 = new SolidBrush(Color.FromArgb(129, 140, 248)))
                {
                    g.FillEllipse(dot1, 23, 8, 3, 3);
                    g.FillEllipse(dot2, 23, 15, 3, 3);
                }
            }
            return bmp;
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
