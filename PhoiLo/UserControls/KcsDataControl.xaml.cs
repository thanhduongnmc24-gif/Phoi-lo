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
    public class KcsRow {
        public string STT { get; set; } = "";
        public string PhuongThuc { get; set; } = "";
        public string MacPhoi { get; set; } = "";
        public string MeSo { get; set; } = "";
        public string SoCayNap { get; set; } = ""; // Mới thêm
        public string ChieuDai { get; set; } = "";
    }

    public partial class KcsDataControl : UserControl
    {
        private List<KcsRow> _dataList = new List<KcsRow>();
        
        public KcsDataControl() {
            InitializeComponent();
            _dataList = Enumerable.Range(1, 50).Select(i => new KcsRow { STT = i.ToString() }).ToList();
            KcsDataGrid.ItemsSource = _dataList;
        }

        private void KcsDataGrid_PreviewKeyDown(object sender, KeyEventArgs e) {
            if (e.Key == Key.V && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control) {
                string[] lines = Clipboard.GetText().Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                int start = KcsDataGrid.SelectedIndex < 0 ? 0 : KcsDataGrid.SelectedIndex;
                for (int i = 0; i < lines.Length && (start + i) < 50; i++) {
                    string[] c = lines[i].Split('\t');
                    // Cập nhật lại dải copy để bao gồm cột Số cây nạp lò
                    if (c.Length > 0) _dataList[start + i].PhuongThuc = c[0];
                    if (c.Length > 1) _dataList[start + i].MacPhoi = c[1];
                    if (c.Length > 2) _dataList[start + i].MeSo = c[2];
                    if (c.Length > 3) _dataList[start + i].SoCayNap = c[3];
                    if (c.Length > 4) _dataList[start + i].ChieuDai = c[4];
                }
                KcsDataGrid.Items.Refresh(); e.Handled = true;
            }
        }

        private void UpdateFileCount() {
            var cfg = App.Config;
            string dateStr = cfg.CurrentDate.ToString("dd-MM-yyyy");
            string subFolderName = $"Kip-{cfg.CurrentKip}-Ngay-{dateStr}";
            string fullFolderPath = Path.Combine("ExportKCS", subFolderName);
            
            if (!Directory.Exists(fullFolderPath)) {
                TxtFileCount.Text = "Số file trong ngày: 0";
                return;
            }
            TxtFileCount.Text = $"Số file trong ngày: {Directory.GetFiles(fullFolderPath, "*.xlsx").Length}";
        }

        private void BtnExportExcel_Click(object sender, RoutedEventArgs e) {
            try {
                var cfg = App.Config;
                string dateExcelStr = cfg.CurrentDate.ToString("dd/MM/yyyy"); // Định dạng hiển thị trong Excel
                string dateFolderStr = cfg.CurrentDate.ToString("dd-MM-yyyy"); // Định dạng tên thư mục
                string kipStr = cfg.CurrentKip;
                
                string subFolderName = $"Kip-{kipStr}-Ngay-{dateFolderStr}";
                string fullFolderPath = Path.Combine("ExportKCS", subFolderName);
                if (!Directory.Exists(fullFolderPath)) Directory.CreateDirectory(fullFolderPath);
                
                int x = Directory.GetFiles(fullFolderPath, "*.xlsx").Length + 1;
                string fileName = $"{x}-{kipStr}-{dateFolderStr}-{DateTime.Now:HHmm}.xlsx";
                string filePath = Path.Combine(fullFolderPath, fileName);
                
                using (var wb = new XLWorkbook()) {
                    var ws = wb.Worksheets.Add("KCS");
                    
                    // [Suy luận] Viết vào ô A1, sau đó gộp từ A1 đến F1 (6 cột)
                    ws.Cell(1, 1).Value = $"Kíp {kipStr} ngày {dateExcelStr}";
                    ws.Range("A1:F1").Merge();
                    ws.Cell(1, 1).Style.Font.Bold = true;
                    ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    
                    // Bỏ dòng 2, viết tiêu đề vào dòng 3
                    ws.Cell(3, 1).Value = "STT"; 
                    ws.Cell(3, 2).Value = "Phương thức nạp"; 
                    ws.Cell(3, 3).Value = "Mác phôi";
                    ws.Cell(3, 4).Value = "Mẻ số"; 
                    ws.Cell(3, 5).Value = "Số cây nạp lò"; 
                    ws.Cell(3, 6).Value = "Chiều dài";
                    ws.Range("A3:F3").Style.Font.Bold = true;

                    int r = 4;
                    foreach (var item in _dataList.Where(i => !string.IsNullOrEmpty(i.MacPhoi) || !string.IsNullOrEmpty(i.MeSo))) {
                        ws.Cell(r, 1).Value = item.STT; 
                        ws.Cell(r, 2).Value = item.PhuongThuc;
                        ws.Cell(r, 3).Value = item.MacPhoi; 
                        ws.Cell(r, 4).Value = item.MeSo; 
                        ws.Cell(r, 5).Value = item.SoCayNap;
                        ws.Cell(r, 6).Value = item.ChieuDai;
                        r++;
                    }
                    ws.Columns().AdjustToContents(); // Tự động co giãn cột cho vừa chữ
                    wb.SaveAs(filePath);
                }
                MessageBox.Show($"Đã xuất file vào thư mục: ExportKCS\\{subFolderName}"); 
                UpdateFileCount();
            } catch (Exception ex) { MessageBox.Show("Lỗi: " + ex.Message); }
        }
    }
}