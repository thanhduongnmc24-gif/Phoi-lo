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
        public string ChieuDai { get; set; } = "";
    }

    public partial class KcsDataControl : UserControl
    {
        private List<KcsRow> _dataList = new List<KcsRow>();
        public KcsDataControl() {
            InitializeComponent();
            _dataList = Enumerable.Range(1, 50).Select(i => new KcsRow { STT = i.ToString() }).ToList();
            KcsDataGrid.ItemsSource = _dataList;
            UpdateFileCount();
        }

        private void KcsDataGrid_PreviewKeyDown(object sender, KeyEventArgs e) {
            if (e.Key == Key.V && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control) {
                string[] lines = Clipboard.GetText().Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                int start = KcsDataGrid.SelectedIndex < 0 ? 0 : KcsDataGrid.SelectedIndex;
                for (int i = 0; i < lines.Length && (start + i) < 50; i++) {
                    string[] c = lines[i].Split('\t');
                    if (c.Length > 0) _dataList[start + i].PhuongThuc = c[0];
                    if (c.Length > 1) _dataList[start + i].MacPhoi = c[1];
                    if (c.Length > 2) _dataList[start + i].MeSo = c[2];
                    if (c.Length > 3) _dataList[start + i].ChieuDai = c[3];
                }
                KcsDataGrid.Items.Refresh(); e.Handled = true;
            }
        }

        private void UpdateFileCount() {
            if (!Directory.Exists("ExportKCS")) Directory.CreateDirectory("ExportKCS");
            string d = App.Config.CurrentDate.ToString("ddMMyyyy");
            TxtFileCount.Text = $"Số file trong ngày: {Directory.GetFiles("ExportKCS", $"*-*-{d}-*.xlsx").Length}";
        }

        private void BtnExportExcel_Click(object sender, RoutedEventArgs e) {
            try {
                var cfg = App.Config;
                string d = cfg.CurrentDate.ToString("ddMMyyyy");
                int x = Directory.GetFiles("ExportKCS", $"*-*-{d}-*.xlsx").Length + 1;
                string fileName = $"{x}-{cfg.CurrentKip}-{d}-{DateTime.Now:HHmm}.xlsx";
                
                using (var wb = new XLWorkbook()) {
                    var ws = wb.Worksheets.Add("KCS");
                    ws.Cell(1, 1).Value = "THÔNG TIN: " + cfg.CurrentKip + " | " + cfg.CurrentToTruong + " | " + cfg.CurrentVanHanh;
                    ws.Cell(3, 1).Value = "STT"; ws.Cell(3, 2).Value = "Phương thức"; ws.Cell(3, 3).Value = "Mác phôi";
                    ws.Cell(3, 4).Value = "Mẻ số"; ws.Cell(3, 5).Value = "Chiều dài";
                    int r = 4;
                    foreach (var item in _dataList.Where(i => !string.IsNullOrEmpty(i.MacPhoi))) {
                        ws.Cell(r, 1).Value = item.STT; ws.Cell(r, 2).Value = item.PhuongThuc;
                        ws.Cell(r, 3).Value = item.MacPhoi; ws.Cell(r, 4).Value = item.MeSo; ws.Cell(r, 5).Value = item.ChieuDai;
                        r++;
                    }
                    ws.Columns().AdjustToContents();
                    wb.SaveAs(Path.Combine("ExportKCS", fileName));
                }
                MessageBox.Show("Đã xuất file: " + fileName); UpdateFileCount();
            } catch (Exception ex) { MessageBox.Show("Lỗi: " + ex.Message); }
        }
    }
}