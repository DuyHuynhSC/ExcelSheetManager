using System;

namespace ExcelSheetManager.Models
{
    public class SheetItem
    {
        public string Name { get; set; } = string.Empty;
        public string ParentWorkbookName { get; set; } = string.Empty;
        public object SheetRef { get; set; } = default!;
        public string TabColorHex { get; set; } = "#3B82F6"; // Default accent color if no tab color
        public bool HasCustomTabColor { get; set; }
        public bool IsVisible { get; set; } = true;
        public bool IsActive { get; set; }
        public string SheetType { get; set; } = "Worksheet";
    }
}
