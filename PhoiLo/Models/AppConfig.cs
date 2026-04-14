using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace PhoiLo.Models
{
    public class AppConfig : INotifyPropertyChanged
    {
        private string _clientId = "";
        private string _clientSecret = "";
        private string _sheetId = "";
        private string _range = "Phoi!A11:J410";
        
        // --- CÁC THÔNG SỐ GIAO DIỆN MỚI ---
        private double _columnWidth = 100;
        private double _tableFontSize = 14;
        private double _menuFontSize = 16;
        private string _tableFontFamily = "Segoe UI";
        private string _menuFontFamily = "Segoe UI";

        public string ClientId { get => _clientId; set { _clientId = value; OnPropertyChanged(); } }
        public string ClientSecret { get => _clientSecret; set { _clientSecret = value; OnPropertyChanged(); } }
        public string SheetId { get => _sheetId; set { _sheetId = value; OnPropertyChanged(); } }
        public string Range { get => _range; set { _range = value; OnPropertyChanged(); } }

        public double ColumnWidth { get => _columnWidth; set { _columnWidth = value; OnPropertyChanged(); } }
        public double TableFontSize { get => _tableFontSize; set { _tableFontSize = value; OnPropertyChanged(); } }
        public double MenuFontSize { get => _menuFontSize; set { _menuFontSize = value; OnPropertyChanged(); } }
        public string TableFontFamily { get => _tableFontFamily; set { _tableFontFamily = value; OnPropertyChanged(); } }
        public string MenuFontFamily { get => _menuFontFamily; set { _menuFontFamily = value; OnPropertyChanged(); } }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}