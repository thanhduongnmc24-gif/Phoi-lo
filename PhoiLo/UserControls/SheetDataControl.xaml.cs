using System;
using System.Collections.Generic;
using System.Data;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Util.Store;

namespace PhoiLo.UserControls
{
    public partial class SheetDataControl : UserControl
    {
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "PhoiLo App";

        public SheetDataControl()
        {
            InitializeComponent();
        }

        private async void BtnConnect_Click(object sender, RoutedEventArgs e)
        {
            string clientId = TxtClientId.Text.Trim();
            string clientSecret = TxtClientSecret.Text.Trim();
            string spreadsheetId = TxtSheetId.Text.Trim();
            string range = TxtRange.Text.Trim();

            if (string.IsNullOrEmpty(clientId) || string.IsNullOrEmpty(clientSecret) || string.IsNullOrEmpty(spreadsheetId))
            {
                MessageBox.Show("Anh hai vui lòng điền đầy đủ Client ID, Client Secret và Sheet ID nha!", "Thiếu thông tin");
                return;
            }

            try
            {
                // Cấu hình thông tin xác thực
                ClientSecrets secrets = new ClientSecrets
                {
                    ClientId = clientId,
                    ClientSecret = clientSecret
                };

                // [Suy luận] Khi ứng dụng chạy trên Windows, lệnh này sẽ mở trình duyệt web lên yêu cầu anh hai đăng nhập Google
                UserCredential credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore("PhoiLo.GoogleAuth.Store"));

                // Khởi tạo dịch vụ Google Sheet
                var service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                // Gửi yêu cầu lấy dữ liệu
                SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                var response = await request.ExecuteAsync();
                IList<IList<object>> values = response.Values;

                if (values != null && values.Count > 0)
                {
                    // Chuyển đổi dữ liệu thành DataTable
                    DataTable dt = new DataTable();
                    
                    // Dòng đầu tiên trong Sheet được dùng làm Header (Tiêu đề cột)
                    var headers = values[0];
                    foreach (var header in headers)
                    {
                        dt.Columns.Add(header.ToString());
                    }

                    // Các dòng tiếp theo là dữ liệu
                    for (int i = 1; i < values.Count; i++)
                    {
                        var row = dt.NewRow();
                        var rowData = values[i];
                        for (int j = 0; j < headers.Count; j++)
                        {
                            if (j < rowData.Count)
                            {
                                row[j] = rowData[j]?.ToString() ?? "";
                            }
                        }
                        dt.Rows.Add(row);
                    }

                    // Đưa dữ liệu lên DataGrid
                    SheetDataGrid.ItemsSource = dt.DefaultView;
                    MessageBox.Show("Lấy dữ liệu thành công rồi anh hai!", "Thông báo");
                }
                else
                {
                    MessageBox.Show("Không tìm thấy dữ liệu trong bảng tính.", "Thông báo");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi xảy ra: " + ex.Message, "Lỗi");
            }
        }
    }
}