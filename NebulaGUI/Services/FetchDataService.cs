using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NebulaGUI.Models;

namespace NebulaGUI.Services
{
    public class FetchDataService
    {    private static readonly object fileLock = new object();

        public IEnumerable<Datas> FetchData(string filePath)
        {
            bool isFirstLine = true;
            var records = new List<Datas>();

            try
            {
                lock (fileLock)
                {
                    using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    using (StreamReader sr = new StreamReader(fs))
                    using (TextFieldParser parser = new TextFieldParser(sr))
                    {
                        int index = 0;
                        parser.TextFieldType = FieldType.Delimited;
                        parser.SetDelimiters(",");
                        while (!parser.EndOfData)
                        {
                            string[] fields = parser.ReadFields();
                            if (isFirstLine)
                            {
                                isFirstLine = false;
                                continue;
                            }
                            Datas record = new Datas
                            {
                            PaketNo = ParseDouble(fields[0]),
                            UyduStatusu = ParseDouble(fields[1]),
                            HataKodu = ParseDouble(fields[2]),
                            GondermeSaati = ParseDouble(fields[3]),
                            Basinc1 = ParseDouble(fields[4]),
                            Basinc2 = ParseDouble(fields[5]),
                            Yukseklik1 = ParseDouble(fields[6]),
                            Yukseklik2 = ParseDouble(fields[7]),
                            IrtifaFarki = ParseDouble(fields[8]),
                            InisHizi = ParseDouble(fields[9]),
                            Sicaklik = ParseDouble(fields[10]),
                            PilGerilimi = ParseDouble(fields[11]),
                            GpsLatitude = ParseDouble(fields[12]),
                            GpsLongitude = ParseDouble(fields[13]),
                            GpsAltitude = ParseDouble(fields[14]),
                            Pitch = ParseDouble(fields[15]),
                            Roll = ParseDouble(fields[16]),
                            Yaw = ParseDouble(fields[17]),
                            IoTData = ParseDouble(fields[18]),
                            TakimNo = ParseDouble(fields[19]),
                            RHRH = string.IsNullOrWhiteSpace(fields[20]) ? "0" : fields[20],
                            Ayrilma = string.IsNullOrWhiteSpace(fields[21]) ? "0" : fields[21]

                            };
                            records.Add(record);
                        }
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

            return records;
        }

        private double ParseDouble(string value)
        {
            double result;
            if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out result))
            {
                return result;
            }
            return 0;
        }
    }
}