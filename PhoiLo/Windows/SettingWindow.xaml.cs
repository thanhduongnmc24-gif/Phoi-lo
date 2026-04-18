using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using PhoiLo.Helpers;

namespace PhoiLo.Windows
{
    public partial class SettingWindow : Window
    {
        public SettingWindow()
        {
            InitializeComponent();
            this.DataContext = App.Config;
            // Tèo đã xóa lệnh EnableWidthAutoSave ở đây cho nó khỏi phàn nàn thiếu tên
        }

        // [Suy luận] Hàm này được gọi khi giao diện thực sự đã vẽ xong cái DataGrid
        private void StaffGrid_Loaded(object sender, RoutedEventArgs e)
        {
            if (sender is DataGrid grid)
            {
                DataGridHelper.EnableWidthAutoSave(grid, "Staff");
            }
        }

        private void StaffGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (sender is DataGrid grid)
            {
                DataGridHelper.HandleExcelActions(grid, e);
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            App.Config.SaveToFile();
            this.Close();
        }
    }
}