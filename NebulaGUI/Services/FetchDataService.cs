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
    {
        public IEnumerable<Datas> FetchData(string filePath)
        {
            bool isFirstLine = true;
            var records = new List<Datas>();
            using (FileStream fs = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (StreamReader sr = new StreamReader(fs))
            using (TextFieldParser parser = new TextFieldParser(sr))
            {
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
                        RHRH = fields[20]
                    };
                    records.Add(record);
                }
            }
            return records;
        }

        private double ParseDouble(string input)
        {
            if (double.TryParse(input, NumberStyles.Any, CultureInfo.InvariantCulture, out double result))
            {
                return result;
            }
            else
            {
                // Hatalı giriş durumunda uygun bir işlem yapın (örneğin, loglama)
                Console.WriteLine($"Hatalı giriş dizesi: {input}");
                return 0.0; // veya uygun bir varsayılan değer
            }
        }



    }
}