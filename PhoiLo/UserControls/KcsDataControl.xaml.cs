using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ClosedXML.Excel;

namespace PhoiLo.UserControls
{
    public class KcsRow
    {
        public string STT { get; set; } = "";
        public string PhuongThuc { get; set; } = "";
        public string MacPhoi { get; set; } = "";
        public string MeSo { get; set; } = "";
        public string ChieuDai { get; set; } = "";
        public string Kip { get; set; } = "";
        public string VanHanh { get; set; } = "";
        public string ToTruong { get; set; } = "";
    }

    public partial class KcsDataControl : UserControl
    {
        private List<KcsRow> _dataList = new List<KcsRow>();

        public KcsDataControl()
        {
            InitializeComponent();
            DpDate.SelectedDate = DateTime.Now;
            InitDataGrid();
            UpdateFileCount();
        }

        private void InitDataGrid()
        {
            // [Suy luận] Khởi tạo sẵn 50 dòng để copy paste thả ga
            _dataList = Enumerable.Range(1, 50).Select(i => new KcsRow { STT = i.ToString() }).ToList();
            KcsDataGrid.ItemsSource = _dataList;
        }

        private void CbKip_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CbKip.SelectedItem is ComboBoxItem selectedItem)
            {
                string fullKip = selectedItem.Content.ToString()!;
                string group = fullKip.Substring(fullKip.Length - 1); // Cắt lấy đuôi A, B, C

                var staff = App.Config.StaffList;
                CbToTruong.ItemsSource = staff.Where(s => s.Kip.ToUpper() == group && s.ChứcVu.Contains("Tổ trưởng")).ToList();
                CbVanHanh.ItemsSource = staff.Where(s => s.Kip.ToUpper() == group && s.ChứcVu.Contains("Vận hành")).ToList();
                
                if (CbToTruong.Items.Count > 0) CbToTruong.SelectedIndex = 0;
                if (CbVanHanh.Items.Count > 0) CbVanHanh.SelectedIndex = 0;

                UpdateDataGridKip();
            }
        }

        private void CbNhanSu_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateDataGridKip();
        }

        private void UpdateDataGridKip()
        {
            // Cập nhật lại cột Kíp, Nhân sự cho toàn bộ 50 dòng
            string kip = (CbKip.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "";
            string toTruong = CbToTruong.Text;
            string vanHanh = CbVanHanh.Text;

            foreach (var row in _dataList)
            {
                row.Kip = kip;
                row.ToTruong = toTruong;
                row.VanHanh = vanHanh;
            }
            if (KcsDataGrid != null) KcsDataGrid.Items.Refresh();
        }

        private void DpDate_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateFileCount();
        }

        private void KcsDataGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            // [Suy luận] Bắt phím Ctrl + V để lấy dữ liệu từ Clipboard
            if (e.Key == Key.V && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                string clipboardText = Clipboard.GetText();
                string[] lines = clipboardText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                
                int startRow = KcsDataGrid.SelectedIndex;
                if (startRow < 0) startRow = 0;

                for (int i = 0; i < lines.Length && (startRow + i) < _dataList.Count; i++)
                {
                    string[] cells = lines[i].Split('\t');
                    var row = _dataList[startRow + i];
                    
                    if (cells.Length > 0) row.PhuongThuc = cells[0];
                    if (cells.Length > 1) row.MacPhoi = cells[1];
                    if (cells.Length > 2) row.MeSo = cells[2];
                    if (cells.Length > 3) row.ChieuDai = cells[3];
                }
                KcsDataGrid.Items.Refresh();
                e.Handled = true;
            }
        }

        private void UpdateFileCount()
        {
            if (!Directory.Exists("ExportKCS")) Directory.CreateDirectory("ExportKCS");
            string datePattern = DpDate.SelectedDate?.ToString("dd-MM-yyyy") ?? DateTime.Now.ToString("dd-MM-yyyy");
            int count = Directory.GetFiles("ExportKCS", $"*-*-{datePattern}-*.xlsx").Length;
            if (TxtFileCount != null) TxtFileCount.Text = $"Số file trong ngày: {count}";
        }

        private void BtnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string dateStr = DpDate.SelectedDate?.ToString("dd-MM-yyyy") ?? DateTime.Now.ToString("dd-MM-yyyy");
                string timeStr = DateTime.Now.ToString("HH-mm");
                string kipStr = (CbKip.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "ChuaChonKip";
                
                if (!Directory.Exists("ExportKCS")) Directory.CreateDirectory("ExportKCS");
                
                int x = Directory.GetFiles("ExportKCS", $"*-*-{dateStr}-*.xlsx").Length + 1;
                
                // Định dạng tên theo ý anh hai: x-kíp-ngày-tháng-năm-giờ-phút.xlsx
                string fileName = $"{x}-{kipStr}-{dateStr}-{timeStr}.xlsx";
                string filePath = Path.Combine("ExportKCS", fileName);

                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("KCS");
                    
                    ws.Cell(1, 1).Value = "BIÊN BẢN GỞI PHÔI KCS";
                    ws.Cell(2, 1).Value = $"Ngày: {DpDate.SelectedDate?.ToString("dd/MM/yyyy")}";
                    ws.Cell(2, 2).Value = $"Kíp: {kipStr}";
                    ws.Cell(3, 1).Value = $"Tổ trưởng: {CbToTruong.Text}";
                    ws.Cell(3, 2).Value = $"Vận hành: {CbVanHanh.Text}";

                    string[] headers = { "STT", "Phương thức nạp", "Mác phôi", "Mẻ số", "Chiều dài", "Kíp", "Vận hành", "Tổ trưởng" };
                    for (int i = 0; i < headers.Length; i++) ws.Cell(5, i + 1).Value = headers[i];

                    int rowIdx = 6;
                    // Lọc những dòng nào có người dùng nhập liệu (Mác phôi hoặc Mẻ số) mới xuất ra
                    foreach (var item in _dataList.Where(d => !string.IsNullOrEmpty(d.MacPhoi) || !string.IsNullOrEmpty(d.MeSo)))
                    {
                        ws.Cell(rowIdx, 1).Value = item.STT;
                        ws.Cell(rowIdx, 2).Value = item.PhuongThuc;
                        ws.Cell(rowIdx, 3).Value = item.MacPhoi;
                        ws.Cell(rowIdx, 4).Value = item.MeSo;
                        ws.Cell(rowIdx, 5).Value = item.ChieuDai;
                        ws.Cell(rowIdx, 6).Value = kipStr;
                        ws.Cell(rowIdx, 7).Value = CbVanHanh.Text;
                        ws.Cell(rowIdx, 8).Value = CbToTruong.Text;
                        rowIdx++;
                    }

                    ws.Columns().AdjustToContents();
                    workbook.SaveAs(filePath);
                }

                MessageBox.Show($"Đã xuất file thành công!\nThư mục: ExportKCS\nTên file: {fileName}", "Hoàn tất", MessageBoxButton.OK, MessageBoxImage.Information);
                UpdateFileCount();
            }
            catch (Exception ex) { MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error); }
        }
    }
}