using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using NebulaGUI.Models;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace NebulaGUI.Services
{
    public class FetchDataService
    {
        private static readonly SemaphoreSlim semaphore = new SemaphoreSlim(1, 1);

        public async Task<IEnumerable<Datas>> FetchDataAsync(string filePath)
        {
            bool isFirstLine = true;
            var records = new List<Datas>();

            await semaphore.WaitAsync(); // Asenkron kilitleme
            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (StreamReader sr = new StreamReader(fs))
                using (TextFieldParser parser = new TextFieldParser(sr))
                {
                    parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters(","); // Delimiter olarak virgül ayarlandı
                    while (!parser.EndOfData)
                    {
                        string[] fields = parser.ReadFields();
                        if (isFirstLine)
                        {
                            isFirstLine = false;
                            continue;
                        }

                        // Ensure the array has enough elements
                        Datas record = ParseFieldsWithDefaults(fields);
                        records.Add(record);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Hata: {ex.Message}");
            }
            finally
            {
                semaphore.Release(); // Kilidi serbest bırak
            }

            return records;
        }

        private Datas ParseFieldsWithDefaults(string[] fields)
        {
            var data = new Datas
            {
                PaketNo = fields.Length > 0 ? ParseDouble(fields[0]) : 0,
                UyduStatusu = fields.Length > 1 ? ParseDouble(fields[1]) : 0,
                HataKodu = fields.Length > 2 ? ParseDouble(fields[2]) : 0,
                GondermeSaati = fields.Length > 3 ? ParseDouble(fields[3]) : 0,
                Basinc1 = fields.Length > 4 ? ParseDouble(fields[4]) : 0,
                Basinc2 = fields.Length > 5 ? ParseDouble(fields[5]) : 0,
                Yukseklik1 = fields.Length > 6 ? ParseDouble(fields[6]) : 0,
                Yukseklik2 = fields.Length > 7 ? ParseDouble(fields[7]) : 0,
                IrtifaFarki = fields.Length > 8 ? ParseDouble(fields[8]) : 0,
                InisHizi = fields.Length > 9 ? ParseDouble(fields[9]) : 0,
                Sicaklik = fields.Length > 10 ? ParseDouble(fields[10]) : 0,
                PilGerilimi = fields.Length > 11 ? ParseDouble(fields[11]) : 0,
                GpsLatitude = fields.Length > 12 ? ParseDouble(fields[12]) : 0,
                GpsLongitude = fields.Length > 13 ? ParseDouble(fields[13]) : 0,
                GpsAltitude = fields.Length > 14 ? ParseDouble(fields[14]) : 0,
                Pitch = fields.Length > 15 ? ParseDouble(fields[15]) : 0,
                Roll = fields.Length > 16 ? ParseDouble(fields[16]) : 0,
                Yaw = fields.Length > 17 ? ParseDouble(fields[17]) : 0,
                IoTData = fields.Length > 18 ? ParseDouble(fields[18]) : 0,
                TakimNo = fields.Length > 19 ? ParseDouble(fields[19]) : 0,
                RHRH = fields.Length > 20 && !string.IsNullOrWhiteSpace(fields[20]) ? fields[20] : "0",
                Ayrilma = fields.Length > 21 && !string.IsNullOrWhiteSpace(fields[21]) ? fields[21] : "0"
            };

            return data;
        }

        private double ParseDouble(string value)
        {
            if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double result))
            {
                return result;
            }
            return 0;
        }

        public async Task UpdateExcelFileAsync(string filePath, string komutText, string ayrilmakomutText)
        {
            await semaphore.WaitAsync(); // Asenkron kilitleme
            try
            {
                await Task.Run(() =>
                {
                    var excelApp = new Application();
                    Workbook workbook = null;
                    Worksheet worksheet = null;

                    try
                    {
                        excelApp.DisplayAlerts = false;
                        workbook = excelApp.Workbooks.Open(filePath, ReadOnly: false, Editable: true);
                        worksheet = workbook.Sheets[1];

                        Range aColumn = worksheet.Columns["A"];
                        int rowCount = aColumn.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;

                        for (int i = 1; i <= rowCount; i++)
                        {
                            Range vcell = worksheet.Cells[i, "U"];
                            if (vcell.Value == null)
                            {
                                vcell.Value = komutText;
                            }

                            Range wcell = worksheet.Cells[i, "V"];
                            if (wcell.Value == null)
                            {
                                wcell.Value = ayrilmakomutText;
                            }
                        }

                        workbook.Save();
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
                });
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Hata: {ex.Message}");
            }
            finally
            {
                semaphore.Release(); // Kilidi serbest bırak
            }
        }

        public async Task<(double latitude, double longitude, double altitude, double roll, double yaw, double pitch)> GetLastGpsAndOrientationDataAsync(string filePath)
        {
            await semaphore.WaitAsync();
            try
            {
                return await Task.Run(() =>
                {
                    var excelApp = new Application();
                    Workbook workbook = null;
                    Worksheet worksheet = null;
                    double latitude = 0, longitude = 0, altitude = 0;
                    double roll = 0, yaw = 0, pitch = 0;

                    try
                    {
                        excelApp.DisplayAlerts = false;
                        workbook = excelApp.Workbooks.Open(filePath);
                        worksheet = workbook.Sheets[1];

                        Range aColumn = worksheet.Columns["A"];
                        int rowCount = aColumn.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;

                        if (rowCount > 1)
                        {
                            Range latCell = worksheet.Cells[rowCount, "M"];
                            Range lonCell = worksheet.Cells[rowCount, "N"];
                            Range altCell = worksheet.Cells[rowCount, "O"];
                            Range rollCell = worksheet.Cells[rowCount, "P"];
                            Range yawCell = worksheet.Cells[rowCount, "Q"];
                            Range pitchCell = worksheet.Cells[rowCount, "R"];

                            latitude = latCell.Value != null ? latCell.Value : 0;
                            longitude = lonCell.Value != null ? lonCell.Value : 0;
                            altitude = altCell.Value != null ? altCell.Value : 0;
                            roll = rollCell.Value != null ? rollCell.Value : 0;
                            yaw = yawCell.Value != null ? yawCell.Value : 0;
                            pitch = pitchCell.Value != null ? pitchCell.Value : 0;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Windows.MessageBox.Show($"Hata: {ex.Message}");
                    }
                    finally
                    {
                        if (workbook != null)
                        {
                            workbook.Close(false);
                            Marshal.ReleaseComObject(workbook);
                        }
                        if (excelApp != null)
                        {
                            excelApp.Quit();
                            Marshal.ReleaseComObject(excelApp);
                        }
                    }

                    return (latitude, longitude, altitude, roll, yaw, pitch);
                });
            }
            finally
            {
                semaphore.Release();
            }
        }
    }
}

