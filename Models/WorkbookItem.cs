using System;

namespace ExcelSheetManager.Models
{
    public class WorkbookItem
    {
        public string Name { get; set; } = string.Empty;
        public string FullPath { get; set; } = string.Empty;
        public object WorkbookRef { get; set; } = default!;
        public int SheetCount { get; set; }
        public bool IsActive { get; set; }

        public string DisplaySubtitle => string.IsNullOrEmpty(FullPath) ? "Unsaved Workbook" : FullPath;
    }
}
