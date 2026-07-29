using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ExcelSheetManager.Models;
using ExcelSheetManager.ViewModels;

namespace ExcelSheetManager.Views
{
    public partial class TaskPaneView : System.Windows.Controls.UserControl
    {
        public TaskPaneView()
        {
            InitializeComponent();
            DataContext = new MainViewModel();
        }

        public MainViewModel? ViewModel => DataContext as MainViewModel;

        private void WorkbooksListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ViewModel?.SelectedWorkbook != null)
            {
                ViewModel.ActivateWorkbook(ViewModel.SelectedWorkbook);
            }
        }

        private void WorkbookItem_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is FrameworkElement element && element.DataContext is WorkbookItem workbookItem)
            {
                ViewModel?.ActivateWorkbook(workbookItem);
            }
        }

        private void WorkbooksListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ViewModel?.SelectedWorkbook != null)
            {
                ViewModel.ActivateWorkbook(ViewModel.SelectedWorkbook);
            }
        }

        private void SheetsListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ViewModel?.SelectedSheet != null)
            {
                ViewModel.ActivateSheet(ViewModel.SelectedSheet);
            }
        }

        private void SheetItem_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is FrameworkElement element && element.DataContext is SheetItem sheetItem)
            {
                ViewModel?.ActivateSheet(sheetItem);
            }
        }
    }
}
