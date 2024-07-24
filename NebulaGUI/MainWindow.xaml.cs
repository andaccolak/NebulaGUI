using AForge.Video;
using AForge.Video.DirectShow;
using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;
using System.IO;
using System.Drawing.Imaging;
using System.Drawing;
using Application = Microsoft.Office.Interop.Excel.Application;
using FilterCategory = AForge.Video.DirectShow.FilterCategory;

namespace NebulaGUI
{
    public partial class MainWindow : System.Windows.Window
    {

        private static readonly object fileLock = new object();

        private DispatcherTimer timer;
        private FilterInfoCollection videoDevices;
        private VideoCaptureDevice videoSource;

        public MainWindow()
        {
            InitializeComponent();
            InitializeWebcam();
            InitializeTimer();
        }

        private void InitializeWebcam()
        {
            videoDevices = new FilterInfoCollection(FilterCategory.VideoInputDevice);
            if (videoDevices.Count == 0)
            {
                MessageBox.Show("No video devices found");
                return;
            }

            videoSource = new VideoCaptureDevice(videoDevices[1].MonikerString);
            videoSource.NewFrame += new NewFrameEventHandler(video_NewFrame);
            videoSource.Start();
        }

        private void video_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            try
            {
                BitmapImage bi;
                using (var bitmap = (System.Drawing.Bitmap)eventArgs.Frame.Clone())
                {
                    bi = BitmapToBitmapImage(bitmap);
                }

                bi.Freeze();
                Dispatcher.BeginInvoke(new System.Action(() => WebcamImage.Source = bi));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private BitmapImage BitmapToBitmapImage(System.Drawing.Bitmap bitmap)
        {
            using (var memory = new MemoryStream())
            {
                bitmap.Save(memory, ImageFormat.Bmp);
                memory.Position = 0;
                var bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = memory;
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();
                return bitmapImage;
            }
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            if (videoSource.IsRunning)
            {
                videoSource.SignalToStop();
                videoSource.WaitForStop();
            }
            base.OnClosing(e);
        }

        private void InitializeTimer()
        {
            timer = new DispatcherTimer();
            timer.Interval = TimeSpan.FromSeconds(2);
            timer.Tick += Timer_Tick;
            timer.Start();
        }

        private async void Timer_Tick(object sender, EventArgs e)
        {
            await Task.Run(() => UpdateExcelFile());
        }

        private void UpdateExcelFile()
        {
            string filePath = "C:\\Users\\colak\\OneDrive\\Belgeler\\Kitap1.csv";
            string newFilePath = "C:\\Users\\colak\\OneDrive\\Masaüstü\\WPF\\Kitap1_updated.csv";
            string textBoxValue = Dispatcher.Invoke(() => komut.Text);
            string textBoxValue1 = Dispatcher.Invoke(() => komut_Copy.Text);

           
            try
            {
                lock (fileLock)
                {
                    var excelApp = new Application();
                    Workbook workbook = null;
                    Worksheet worksheet = null;

                    try
                    {
                        excelApp.DisplayAlerts = false;
                        workbook = excelApp.Workbooks.Open(filePath);
                        worksheet = workbook.Sheets[1];

                        Range aColumn = worksheet.Columns["A"];
                        int rowCount = aColumn.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;

                        for (int i = 1; i <= rowCount; i++)
                        {
                            Range vcell = worksheet.Cells[i, "V"];
                            if (vcell.Value == null)
                            {
                                vcell.Value = textBoxValue;
                            }

                            Range wcell = worksheet.Cells[i, "W"];
                            if (wcell.Value == null)
                            {
                                wcell.Value = textBoxValue1;
                            }
                        }

                        workbook.SaveAs(newFilePath);
                    }
                    finally
                    {
                        if (workbook != null)
                        {
                            workbook.Close(false);
                            Marshal.ReleaseComObject(workbook);
                        }
                        excelApp.Quit();
                        Marshal.ReleaseComObject(excelApp);
                    }
                }
            }
            catch (IOException ioEx)
            {
                System.Windows.MessageBox.Show($"Dosya erişim hatası: {ioEx.Message}");
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Hata: {ex.Message}");
            }
        }
    

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Console.WriteLine("deneme");
        }

        private void ayrilma_Click(object sender, RoutedEventArgs e)
        {
            komut_Copy.Text = "1";
        }
    }
}
