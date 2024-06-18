
using Microsoft.Office.Interop.Excel;
using System;

using System.Runtime.InteropServices;

using System.Windows;

using System.Windows.Threading;

namespace NebulaGUI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        private DispatcherTimer timer;

        public MainWindow()
        {
            InitializeComponent();
            InitializeTimer();
        }

        private void InitializeTimer()
        {
            timer = new DispatcherTimer();
            timer.Interval = TimeSpan.FromSeconds(1);
            timer.Tick += Timer_Tick;
            timer.Start();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            UpdateExcelFile();
        }

        private void UpdateExcelFile()
        {
            string filePath = "C:\\Users\\colak\\OneDrive\\Masaüstü\\WPF\\Kitap1.csv";
            string textBoxValue = komut.Text;

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excelApp.Workbooks.Open(filePath);
            Worksheet worksheet = workbook.Sheets[1];

            try
            {
                Range aColumn = worksheet.Columns["A"];
                int rowCount = aColumn.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;

                for (int i = 1; i <= rowCount; i++)
                {
                    Range cell = worksheet.Cells[i, "V"];
                    if (cell.Value == null)
                    {
                        cell.Value = textBoxValue;
                    }
                }

                workbook.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Hata: {ex.Message}");
            }
            finally
            {
                workbook.Close(false);
                Marshal.ReleaseComObject(workbook);
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
        }
    }

}
