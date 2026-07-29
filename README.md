# 📊 Excel Sheet & File Manager Add-in

![.NET 8.0](https://img.shields.io/badge/.NET-8.0-512BD4?style=flat&logo=dotnet)
![Excel-DNA](https://img.shields.io/badge/Excel--DNA-1.9.0-217346?style=flat&logo=microsoft-excel)
![WPF](https://img.shields.io/badge/UI-WPF%20%26%20WinForms-0078D4?style=flat)
![Platform](https://img.shields.io/badge/Platform-Windows-blue?style=flat&logo=windows)

**Excel Sheet & File Manager** là một Excel Add-in (Custom Taskpane) hiện đại được xây dựng bằng C# .NET 8.0 và [Excel-DNA](https://excel-dna.net/). Add-in giúp người dùng quản lý, di chuyển nhanh giữa các file Excel đang mở và các sheet trong từng file một cách trực quan, mượt mà.

---

## ✨ Tính Năng Nổi Bật

- 📁 **Taskpane 2 Vùng Trực Quan (Dual-Region Layout)**:
  - **Vùng 1 (Open Files)**: Tự động tải danh sách tất cả các workbook (`.xlsx`, `.xls`, `.xlsm`, v.v.) đang cùng mở trong Excel.
  - **Vùng 2 (Sheets List)**: Tự động hiển thị danh sách toàn bộ các sheet thuộc file Excel được chọn ở Vùng 1.
- 🎨 **Tô Màu Tab Sheet (Tab Color Indicator)**:
  - Đọc thuộc tính màu sắc tab sheet (`Worksheet.Tab.Color`) từ Excel.
  - Hiển thị dải màu (Tab color badge) chính xác bên cạnh từng tên sheet ở Vùng 2.
- 🎯 **Kích Hoạt Sheet & File Nhanh (Click to Focus)**:
  - **Click chọn sheet**: Excel lập tức chuyển trọng tâm (Focus/Activate) sang sheet đó (tự động kích hoạt file tương ứng nếu đang mở file khác).
  - **Double click file**: Chuyển nhanh sang file Excel được chọn.
- 🔍 **Bộ Lọc Tìm Kiếm Realtime (Live Search Filters)**:
  - Lọc nhanh tên file Excel ở Vùng 1.
  - Lọc nhanh tên sheet ở Vùng 2.
- 🔄 **Đồng Bộ Tự Động (Excel Live Sync & Refresh)**:
  - Đăng ký sẵn Event Hooks với Excel (`WorkbookOpen`, `WorkbookBeforeClose`, `SheetActivate`, `WorkbookNewSheet`). Taskpane tự động cập nhật danh sách ngay khi mở/đóng file hoặc tạo sheet mới.
  - Nút **Refresh** ($\circlearrowright$) trên Taskpane toolbar và Ribbon tiện lợi.
- 🎗️ **Tích Hợp Ribbon Menu**:
  - Thêm tab **Sheet Manager** trên thanh Ribbon Excel để bật/tắt Taskpane hoặc làm mới dữ liệu.

---

## 🛠️ Công Nghệ Sử Dụng

- **Language**: C# (.NET 8.0-windows)
- **Add-in Framework**: Excel-DNA 1.9.0
- **UI Framework**: WPF (Windows Presentation Foundation) kết hợp WinForms `ElementHost` trong `CustomTaskPane`
- **Excel Interop**: `Microsoft.Office.Interop.Excel`

---

## 🚀 Hướng Dẫn Biên Dịch (Build Project)

### Yêu cầu môi trường:
- Windows 10/11
- .NET 8.0 SDK trở lên
- Microsoft Excel (2016 / 2019 / 2021 / Office 365)

### Các bước build:
1. Clone repository về máy:
   ```bash
   git clone https://github.com/your-username/ExcelSheetManager.git
   cd ExcelSheetManager
   ```
2. Build dự án ở chế độ Release:
   ```bash
   dotnet build -c Release
   ```
3. Sau khi build thành công, các file `.xll` độc lập (packed self-contained add-in) sẽ được sinh ra tại:
   `bin/Release/net48/publish/`
   - **`ExcelSheetManager-AddIn64-packed.xll`** (Dành cho Excel 64-bit)
   - **`ExcelSheetManager-AddIn-packed.xll`** (Dành cho Excel 32-bit)

---

## 📖 Hướng Dẫn Cài Đặt Vào Excel

1. Mở **Microsoft Excel**.
2. Vào **File** $\rightarrow$ **Options** $\rightarrow$ **Add-ins**.
3. Tại mục **Manage** (ở phía dưới), chọn **Excel Add-ins** $\rightarrow$ Nhấn **Go...**.
4. Chọn **Browse...** $\rightarrow$ Trỏ tới file **`ExcelSheetManager-AddIn64-packed.xll`** (hoặc bản 32-bit tương ứng).
5. Nhấn **OK** để hoàn tất.
6. Thanh Taskpane **Sheet & File Manager** sẽ tự động hiển thị ở bên phải giao diện Excel.

---

## 📂 Cấu Trúc Dự Án

```text
ExcelSheetManager/
├── AddIn.cs                 # Entry point cho Excel-DNA, quản lý Taskpane & Event listeners
├── RibbonController.cs      # Khởi tạo thanh Ribbon "Sheet Manager" trong Excel
├── ExcelSheetManager.csproj # Cấu hình project .NET 8.0, WPF & Excel-DNA dependencies
├── Helpers/
│   └── ColorHelper.cs       # Chuyển đổi màu OLE BGR từ Excel sang Hex RGB (#RRGGBB)
├── Models/
│   ├── WorkbookItem.cs      # Data model cho File Excel đang mở
│   └── SheetItem.cs         # Data model cho Sheet (Worksheet / Chart)
├── ViewModels/
│   ├── ViewModelBase.cs     # INotifyPropertyChanged base class
│   ├── RelayCommand.cs      # ICommand implementation
│   └── MainViewModel.cs     # State manager & logic lọc, focus sheet/file
└── Views/
    ├── TaskPaneView.xaml    # Giao diện WPF Taskpane (Modern Dark Theme)
    ├── TaskPaneView.xaml.cs # Code-behind cho WPF control
    └── TaskPaneHost.cs      # Container WinForms hosting WPF control vào Excel Taskpane
```

---

## 📝 License

Dự án phát hành theo giấy phép [MIT License](LICENSE).
