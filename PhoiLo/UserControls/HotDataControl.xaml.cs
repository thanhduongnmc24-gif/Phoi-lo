using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Controls;

namespace PhoiLo.UserControls
{
    public partial class HotDataControl : UserControl
    {
        public HotDataControl() => InitializeComponent();

        public void SetData(DataTable sourceTable)
        {
            // [Suy luận] Lọc những dòng có phương thức nạp chứa chữ "nóng"
            var filteredRows = sourceTable.AsEnumerable()
                .Where(r => (r["Phương thức nạp"]?.ToString() ?? "").ToLower().Contains("nóng"))
                .ToList();

            if (filteredRows.Any()) {
                DataTable dtHot = filteredRows.CopyToDataTable();
                HotDataGrid.ItemsSource = dtHot.DefaultView;

                double total = 0;
                var stats = new Dictionary<string, double>();
                foreach (var row in filteredRows) {
                    double n; double.TryParse(row["Số cây nạp lò"]?.ToString(), out n);
                    total += n;
                    string cd = row["Chiều dài"]?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(cd) && n > 0) {
                        if (stats.ContainsKey(cd)) stats[cd] += n;
                        else stats.Add(cd, n);
                    }
                }
                TxtSumHot.Text = total.ToString();
                HotLengthGrid.ItemsSource = stats.Select(x => new { ChieuDai = x.Key, SoLuong = x.Value }).OrderBy(x => x.ChieuDai).ToList();
            } else {
                HotDataGrid.ItemsSource = null; TxtSumHot.Text = "0"; HotLengthGrid.ItemsSource = null;
            }
        }
    }
}