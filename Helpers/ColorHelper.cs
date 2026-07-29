using System;
using Microsoft.Office.Interop.Excel;

namespace ExcelSheetManager.Helpers
{
    public static class ColorHelper
    {
        /// <summary>
        /// Converts Excel Tab OLE Color (BGR) to RGB Hex string "#RRGGBB"
        /// Returns (hex, hasCustomColor)
        /// </summary>
        public static (string Hex, bool HasCustomColor) GetSheetTabColorHex(Worksheet sheet)
        {
            try
            {
                var tab = sheet.Tab;
                // ColorIndex of xlColorIndexNone (-4142) means default/no custom tab color
                if (tab == null || (int)tab.ColorIndex == (int)XlColorIndex.xlColorIndexNone)
                {
                    return ("#94A3B8", false); // Neutral slate color for default tab
                }

                // Excel stores OLE color as 0x00BBGGRR
                double rawColor = (double)tab.Color;
                int colorInt = (int)rawColor;

                int r = colorInt & 0xFF;
                int g = (colorInt >> 8) & 0xFF;
                int b = (colorInt >> 16) & 0xFF;

                string hex = $"#{r:X2}{g:X2}{b:X2}";
                return (hex, true);
            }
            catch
            {
                return ("#94A3B8", false);
            }
        }
    }
}
