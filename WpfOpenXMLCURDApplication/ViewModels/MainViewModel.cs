// ViewModels/MainViewModel.cs
using System;
using System.ComponentModel;
using System.Data;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using WpfOpenXMLCURDApplication.Services;
using WpfOpenXMLCURDApplication.Utilities;

namespace WpfOpenXMLCURDApplication.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private readonly IExcelService _excelService;
        private DataTable _dataTable;
        private string _filePath = "path_to_file.xlsx"; // Path to the Excel file

        // Constructor to inject the ExcelService dependency
        public MainViewModel(IExcelService excelService)
        {
            _excelService = excelService;
            CreateCommand = new RelayCommand(Create);
            ReadCommand = new RelayCommand(Read);
            UpdateCommand = new RelayCommand(Update);
            DeleteCommand = new RelayCommand(Delete);
        }

        // Property to bind DataTable to the UI
        public DataTable DataTable
        {
            get => _dataTable;
            set
            {
                _dataTable = value;
                OnPropertyChanged();
            }
        }

        // Commands for CRUD operations
        public ICommand CreateCommand { get; }
        public ICommand ReadCommand { get; }
        public ICommand UpdateCommand { get; }
        public ICommand DeleteCommand { get; }

        // Method to create an Excel file
        private void Create()
        {
            _excelService.CreateExcelFile(_filePath);
        }

        // Method to read data from an Excel file
        private void Read()
        {
            DataTable = _excelService.ReadExcelFile(_filePath);
        }

        // Method to update a cell in the Excel file
        private void Update()
        {
            _excelService.UpdateCell(_filePath, "Sheet1", "R1C1", "Updated Value");
        }

        // Method to delete a row from the Excel file
        private void Delete()
        {
            _excelService.DeleteRow(_filePath, "Sheet1", 1);
        }

        // INotifyPropertyChanged implementation
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

}
