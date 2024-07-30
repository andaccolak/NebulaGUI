using AForge.Video;
using AForge.Video.DirectShow;
using System;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Microsoft.Office.Interop.Excel;

namespace NebulaGUI
{
    public partial class MainWindow : System.Windows.Window
    {
        private FilterInfoCollection videoDevices;
        private VideoCaptureDevice videoSource;

        public MainWindow()
        {
            InitializeComponent();
            InitializeWebcam();
           
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

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Console.WriteLine("deneme");
        }

        private void ayrilma_Click(object sender, RoutedEventArgs e)
        {
            komut_Copy.Text = "1";
        }
        int buttonCount = 0;
        private void ListUpButton(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {

            buttonCount++;
            if (buttonCount%2==1)
            {
                TempBorder.Visibility = Visibility.Hidden;
                TempText.Visibility = Visibility.Hidden;
                VelocityBorder.Visibility = Visibility.Hidden;
                VelocityText.Visibility = Visibility.Hidden;
                PressureBorder.Visibility = Visibility.Hidden;
                PressureText.Visibility = Visibility.Hidden;
                VoltageBorder.Visibility = Visibility.Hidden;
                VoltageText.Visibility = Visibility.Hidden;
                AltitudeBorder.Visibility = Visibility.Hidden;
                AltitudeText.Visibility = Visibility.Hidden;
                SatBorder.Margin = new Thickness(1438, 16, 10, 608);
                SatText.Margin = new Thickness(1449, 26, 0, 0);
                DataList.Margin = new Thickness(500, 415, 10, 24);
                listButton.Margin = new Thickness(1848, 380, 8, 559);

                ayrilmaCircle.Visibility = Visibility.Visible;
                inisCircle.Visibility = Visibility.Visible;
                kurtarmaCircle.Visibility = Visibility.Visible;
                readyCircle.Visibility = Visibility.Visible;
                yukselmeCircle.Visibility = Visibility.Visible;

                timeBorder.Visibility = Visibility.Visible;
                timeRect.Visibility = Visibility.Visible;
                timeRect2.Visibility = Visibility.Visible;
                Timing.Visibility = Visibility.Visible;

                AyrilmaText.Visibility = Visibility.Visible;
                inisText.Visibility = Visibility.Visible;
                kurtarmaText.Visibility = Visibility.Visible;
                readyText.Visibility = Visibility.Visible;
                yukselmeText.Visibility = Visibility.Visible;
                alarmBorder.Visibility = Visibility.Visible;

            }
            if (buttonCount % 2 == 0)
            {
                TempBorder.Visibility = Visibility.Visible;
                TempText.Visibility = Visibility.Visible;
                VelocityBorder.Visibility = Visibility.Visible;
                VelocityText.Visibility = Visibility.Visible;
                PressureBorder.Visibility = Visibility.Visible;
                PressureText.Visibility = Visibility.Visible;
                VoltageBorder.Visibility = Visibility.Visible;
                VoltageText.Visibility = Visibility.Visible;
                AltitudeBorder.Visibility = Visibility.Visible;
                AltitudeText.Visibility = Visibility.Visible;
                SatBorder.Margin = new Thickness(1438, 391, 10, 269);
                SatText.Margin = new Thickness(1449, 394, 0, 0);
                DataList.Margin = new Thickness(500, 738, 10, 24);
                listButton.Margin = new Thickness(1847, 938, 9, 2);

                ayrilmaCircle.Visibility = Visibility.Hidden;
                inisCircle.Visibility = Visibility.Hidden;
                kurtarmaCircle.Visibility = Visibility.Hidden;
                readyCircle.Visibility = Visibility.Hidden;
                yukselmeCircle.Visibility = Visibility.Hidden;

                timeBorder.Visibility = Visibility.Hidden;
                timeRect.Visibility = Visibility.Hidden;
                timeRect2.Visibility = Visibility.Hidden;
                Timing.Visibility = Visibility.Hidden;

                AyrilmaText.Visibility = Visibility.Hidden;
                inisText.Visibility = Visibility.Hidden;
                kurtarmaText.Visibility = Visibility.Hidden;
                readyText.Visibility = Visibility.Hidden;
                yukselmeText.Visibility = Visibility.Hidden;
                alarmBorder.Visibility = Visibility.Hidden;

            }



        }
    }
}
