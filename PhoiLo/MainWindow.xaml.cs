using System.Windows;
using PhoiLo.UserControls;

namespace PhoiLo
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            // Khởi động là tự động load màn hình Sheet lên Pannel
            MainPanel.Content = new SheetDataControl();
        }

        private void BtnSheet_Click(object sender, RoutedEventArgs e)
        {
            MainPanel.Content = new SheetDataControl();
        }

        private void BtnFeature2_Click(object sender, RoutedEventArgs e)
        {
            // Sau này anh hai tạo thêm UserControl khác thì gọi ở đây
            MessageBox.Show("Tính năng này Tèo đang xây dựng nha anh hai!", "Thông báo");
        }
    }
}