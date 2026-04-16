using System.Linq;
using System.Windows;
using System.Windows.Controls;
using PhoiLo.UserControls;
using PhoiLo.Windows;

namespace PhoiLo
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            MainContentArea.Content = new SheetDataControl();
            LoadStaffToHeader();
        }

        private void CbGlobalKip_SelectionChanged(object sender, SelectionChangedEventArgs e) => LoadStaffToHeader();

        private void LoadStaffToHeader()
        {
            if (CbGlobalKip == null) return;
            string kip = App.Config.CurrentKip;
            if (string.IsNullOrEmpty(kip)) return;
            string group = kip.Substring(kip.Length - 1).ToUpper();

            var staff = App.Config.StaffList;
            CbGlobalToTruong.ItemsSource = staff.Where(s => s.Kip.ToUpper() == group && s.ChứcVu.Contains("Tổ trưởng")).ToList();
            CbGlobalVanHanh.ItemsSource = staff.Where(s => s.Kip.ToUpper() == group && s.ChứcVu.Contains("Vận hành")).ToList();
            
            CbGlobalToTruong.SelectedIndex = 0;
            CbGlobalVanHanh.SelectedIndex = 0;
        }

        private void BtnLoadSheet_Click(object sender, RoutedEventArgs e) => MainContentArea.Content = new SheetDataControl();
        private void BtnLoadKcs_Click(object sender, RoutedEventArgs e) => MainContentArea.Content = new KcsDataControl();
        private void BtnOpenSetting_Click(object sender, RoutedEventArgs e) {
            SettingWindow setWin = new SettingWindow { Owner = this };
            setWin.ShowDialog();
            LoadStaffToHeader(); // Cập nhật lại header sau khi sửa nhân sự
        }
    }
}