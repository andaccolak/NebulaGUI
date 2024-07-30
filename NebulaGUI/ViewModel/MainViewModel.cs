using LiveCharts;
using LiveCharts.Wpf;
using GMap.NET;
using GMap.NET.MapProviders;
using GMap.NET.WindowsPresentation;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Threading;
using HelixToolkit.Wpf;
using System.Windows.Media.Media3D;
using NebulaGUI.Models;
using NebulaGUI.Services;

namespace NebulaGUI.ViewModel
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private DispatcherTimer fetchDataTimer;
        private DispatcherTimer updateExcelTimer;
        private readonly FetchDataService _fetchDataService = new FetchDataService();
        private string dosyaYolu = "C:\\Users\\colak\\Desktop\\WPF\\Kitap123456.csv";

        public ObservableCollection<Datas> AllDatas { get; } = new ObservableCollection<Datas>();
        public ObservableCollection<Datas> AllDatasReversed { get; } = new ObservableCollection<Datas>();

        private string _komut;
        public string Komut
        {
            get { return _komut; }
            set { _komut = value; OnPropertyChanged(nameof(Komut)); }
        }

        private string _ayrilmakomut;
        public string Ayrilmakomut
        {
            get { return _ayrilmakomut; }
            set { _ayrilmakomut = value; OnPropertyChanged(nameof(Ayrilmakomut)); }
        }

        private double _gpsLatitude;
        public double GpsLatitude
        {
            get { return _gpsLatitude; }
            set
            {
                _gpsLatitude = value;
                OnPropertyChanged(nameof(GpsLatitude));
                UpdateGpsPosition();
            }
        }

        private double _gpsLongitude;
        public double GpsLongitude
        {
            get { return _gpsLongitude; }
            set
            {
                _gpsLongitude = value;
                OnPropertyChanged(nameof(GpsLongitude));
                UpdateGpsPosition();
            }
        }

        private double _gpsAltitude;
        public double GpsAltitude
        {
            get { return _gpsAltitude; }
            set { _gpsAltitude = value; OnPropertyChanged(nameof(GpsAltitude)); }
        }

        private PointLatLng _gpsPosition;
        public PointLatLng GpsPosition
        {
            get { return _gpsPosition; }
            set { _gpsPosition = value; OnPropertyChanged(nameof(GpsPosition)); }
        }

        private SeriesCollection _altitudeSeries;
        public SeriesCollection AltitudeSeries
        {
            get { return _altitudeSeries; }
            set { _altitudeSeries = value; OnPropertyChanged(nameof(AltitudeSeries)); }
        }

        private SeriesCollection _temperatureSeries;
        public SeriesCollection TemperatureSeries
        {
            get { return _temperatureSeries; }
            set { _temperatureSeries = value; OnPropertyChanged(nameof(TemperatureSeries)); }
        }

        private SeriesCollection _voltageSeries;
        public SeriesCollection VoltageSeries
        {
            get { return _voltageSeries; }
            set { _voltageSeries = value; OnPropertyChanged(nameof(VoltageSeries)); }
        }

        private SeriesCollection _pressureSeries;
        public SeriesCollection PressureSeries
        {
            get { return _pressureSeries; }
            set { _pressureSeries = value; OnPropertyChanged(nameof(PressureSeries)); }
        }

        private SeriesCollection _speedSeries;
        public SeriesCollection SpeedSeries
        {
            get { return _speedSeries; }
            set { _speedSeries = value; OnPropertyChanged(nameof(SpeedSeries)); }
        }

        private Model3D _model;
        public Model3D Model
        {
            get { return _model; }
            set { _model = value; OnPropertyChanged(nameof(Model)); }
        }

        private Transform3D _modelTransform;
        public Transform3D ModelTransform
        {
            get { return _modelTransform; }
            set { _modelTransform = value; OnPropertyChanged(nameof(ModelTransform)); }
        }

        public MainViewModel()
        {
            fetchDataTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(500)
            };
            fetchDataTimer.Tick += async (s, e) => await VerileriYenileVeGrafikOlusturAsync();
            fetchDataTimer.Start();

            updateExcelTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(1) //0.5 saniyelik veri çekme işleminden sonra 0.5 saniyelik yazma işlemi
            };
            updateExcelTimer.Tick += async (s, e) => await UpdateExcelFileAsync();
            updateExcelTimer.Start();

            LoadModel();
            UpdateGpsAndOrientationData();
        }

        private void LoadModel()
        {
            var loader = new StLReader();
            Model = loader.Read("C:\\Users\\colak\\Desktop\\WPF\\mm.STL");
        }

        private async void UpdateGpsAndOrientationData()
        {
            var data = await _fetchDataService.GetLastGpsAndOrientationDataAsync(dosyaYolu);
            GpsLatitude = data.latitude;
            GpsLongitude = data.longitude;
            GpsAltitude = data.altitude;

            var transformGroup = new Transform3DGroup();
            transformGroup.Children.Add(new RotateTransform3D(new AxisAngleRotation3D(new Vector3D(1, 0, 0), data.pitch)));
            transformGroup.Children.Add(new RotateTransform3D(new AxisAngleRotation3D(new Vector3D(0, 1, 0), data.yaw)));
            transformGroup.Children.Add(new RotateTransform3D(new AxisAngleRotation3D(new Vector3D(0, 0, 1), data.roll)));

            ModelTransform = transformGroup;
        }

        private void UpdateGpsPosition()
        {
            GpsPosition = new PointLatLng(GpsLatitude, GpsLongitude);
        }

        private async Task VerileriYenileVeGrafikOlusturAsync()
        {
            var yuklenenDatas = await _fetchDataService.FetchDataAsync(dosyaYolu);

            if (yuklenenDatas.Any())
            {
                AllDatas.Clear();
                AllDatasReversed.Clear();

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
                        PointGeometry = null,
                        Title = "Altitudes",
                        Values = new ChartValues<double>(altitudes)
                    }
                };

                TemperatureSeries = new SeriesCollection
                {
                    new LineSeries
                    {
                        PointGeometry = null,
                        Title = "Temperature",
                        Values = new ChartValues<double>(temperatures)
                    }
                };

                SpeedSeries = new SeriesCollection
                {
                    new LineSeries
                    {
                        PointGeometry = null,
                        Title = "Speed",
                        Values = new ChartValues<double>(speed)
                    }
                };

                VoltageSeries = new SeriesCollection
                {
                    new LineSeries
                    {
                        PointGeometry = null,
                        Title = "Voltage",
                        Values = new ChartValues<double>(voltage)
                    }
                };

                PressureSeries = new SeriesCollection
                {
                    new LineSeries
                    {
                        PointGeometry = null,
                        Title = "Pressure",
                        Values = new ChartValues<double>(pressure)
                    }
                };
            }
        }

        private async Task UpdateExcelFileAsync()
        {
            await _fetchDataService.UpdateExcelFileAsync(dosyaYolu, Komut, Ayrilmakomut);
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}