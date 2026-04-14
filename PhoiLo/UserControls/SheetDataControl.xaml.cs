using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;

namespace PhoiLo.UserControls
{
    public partial class SheetDataControl : UserControl
    {
        // Khai báo Scope cho phép đọc dữ liệu Sheet
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "PhoiLo App";

        public SheetDataControl()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            LoadGoogleSheetData();
        }

        private void LoadGoogleSheetData()
        {
            try
            {
                // TODO: Anh hai cần bỏ file credentials.json (tải từ Google Cloud Console) vào cùng thư mục chạy của app
                // Hoặc dùng chuỗi ClientId, ClientSecret trực tiếp (nhưng dùng file json là chuẩn bài nhất)
                /* Tạm thời Tèo comment lại đoạn xác thực thật để code không bị lỗi khi chưa có file
                GoogleCredential credential;
                using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                {
                    credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
                }

                var service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                // Thay ID của bảng tính anh hai vào đây
                String spreadsheetId = "ID_BANG_TINH_CUA_ANH_HAI"; 
                String range = "Sheet1!A1:E10"; // Vùng dữ liệu muốn lấy

                SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                var response = request.Execute();
                IList<IList<Object>> values = response.Values;
                */

                // Dữ liệu giả lập (Mock data) để anh hai test giao diện trước khi có API thật
                var mockData = new List<dynamic>
                {
                    new { STT = 1, TenVatTu = "Thép cuộn", SoLuong = 100, TrangThai = "OK" },
                    new { STT = 2, TenVatTu = "Than đá", SoLuong = 500, TrangThai = "Đang chờ" }
                };

                SheetDataGrid.ItemsSource = mockData;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi lấy dữ liệu: " + ex.Message);
            }
        }
    }
}