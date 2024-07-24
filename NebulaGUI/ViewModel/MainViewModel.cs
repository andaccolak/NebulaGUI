using LiveCharts;
using LiveCharts.Wpf;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows.Threading;
using NebulaGUI.Models;
using NebulaGUI.Services;
using System.Threading.Tasks;

namespace NebulaGUI.ViewModel
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private DispatcherTimer timer;
        private string dosyaYolu = "C:\\Users\\colak\\OneDrive\\Belgeler\\Kitap1.csv";


        // Servis ile veri çekme
        private SeriesCollection _altitudeSeries;
        private SeriesCollection _temperatureSeries;
        private SeriesCollection _voltageSeries;
        private SeriesCollection _pressureSeries;
        private SeriesCollection _speedSeries;
        private readonly FetchDataService _denemeService = new FetchDataService();
        public ObservableCollection<Datas> AllDatas { get; } = new ObservableCollection<Datas>();
        public ObservableCollection<Datas> AllDatasReversed { get; } = new ObservableCollection<Datas>();


        public SeriesCollection AltitudeSeries
        {
            get { return _altitudeSeries; }
            set { _altitudeSeries = value; OnPropertyChanged(nameof(AltitudeSeries)); }
        }
        public SeriesCollection SpeedSeries
        {
            get { return _speedSeries; }
            set { _speedSeries = value; OnPropertyChanged(nameof(SpeedSeries)); }
        }

        public SeriesCollection TemperatureSeries
        {
            get { return _temperatureSeries; }
            set { _temperatureSeries = value; OnPropertyChanged(nameof(TemperatureSeries)); }
        }
        public SeriesCollection VoltageSeries
        {
            get { return _voltageSeries; }
            set { _voltageSeries = value; OnPropertyChanged(nameof(VoltageSeries)); }
        }
        public SeriesCollection PressureSeries
        {
            get { return _pressureSeries; }
            set { _pressureSeries = value; OnPropertyChanged(nameof(PressureSeries)); }
        }


        public MainViewModel()
        {
            timer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(2)
            };
            timer.Tick += Timer_Tick;
            timer.Start();
        }

        private async void Timer_Tick(object sender, EventArgs e)
        {
            VerileriYenileVeGrafikOlustur();
        }

        private void VerileriYenileVeGrafikOlustur()
        {

            // Verileri çek, birleştir

            AllDatas.Clear();
            AllDatasReversed.Clear();

            //

            //

            var yuklenenDatas = _denemeService.FetchData(dosyaYolu);
            foreach (var veriData in yuklenenDatas)
            {

                AllDatas.Add(veriData);
            }

            for (int i = AllDatas.Count - 1; i >= 0; i--)
            {
                AllDatasReversed.Add(AllDatas[i]);
            }

            var altitudes = AllDatas.Select(r => r.Yukseklik1).ToList();
            var temperatures = AllDatas.Select(r => r.Sicaklik).ToList();
            var speed = AllDatas.Select(r => r.InisHizi).ToList();
            var voltage = AllDatas.Select(r => r.PilGerilimi).ToList();
            var pressure = AllDatas.Select(r => r.Basinc1).ToList();


            AltitudeSeries = new SeriesCollection
        {
            new LineSeries
            {
                PointGeometry = null, // Noktaları gösterme
                Title = "Altitudes",
                Values = new ChartValues<double>(altitudes)
            }
        };

            TemperatureSeries = new SeriesCollection
            {
                new LineSeries
                {
                    PointGeometry = null, // Noktaları gösterme
                    Title = "Temperature",
                    Values = new ChartValues<double>(temperatures)
                }
            };
            SpeedSeries = new SeriesCollection
            {
                new LineSeries
                {
                    PointGeometry = null, // Noktaları gösterme
                    Title = "hız",
                    Values = new ChartValues<double>(speed)
                }
            };
            VoltageSeries = new SeriesCollection
            {
                new LineSeries
                {
                    PointGeometry = null, // Noktaları gösterme
                    Title = "pil gerilimi",
                    Values = new ChartValues<double>(voltage)
                }
            };
            PressureSeries = new SeriesCollection
            {
                new LineSeries
                {
                    PointGeometry = null, // Noktaları gösterme
                    Title = "basınç",
                    Values = new ChartValues<double>(pressure)
                }
            };
        }


        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}