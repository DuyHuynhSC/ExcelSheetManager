using System;

namespace ExcelSheetManager.Models
{
    public class QuickJumpItem
    {
        public string Title { get; set; } = string.Empty;
        public string Subtitle { get; set; } = string.Empty;
        public string ItemType { get; set; } = "Sheet"; // "Sheet" or "Workbook"
        public string TabColorHex { get; set; } = "#3B82F6";
        public bool HasCustomTabColor { get; set; } = false;
        public bool IsProtected { get; set; } = false;
        public bool IsHidden { get; set; } = false;
        public object? TargetRef { get; set; }
        public string ParentWorkbookName { get; set; } = string.Empty;

        public override string ToString()
        {
            return $"{Title} ({Subtitle})";
        }
    }
}
