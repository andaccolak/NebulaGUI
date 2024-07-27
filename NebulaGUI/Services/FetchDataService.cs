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

            await semaphore.WaitAsync();
            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
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
                        Datas record = ParseFields(fields);
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
                semaphore.Release();
            }

            return records;
        }

        private Datas ParseFields(string[] fields)
        {
            return new Datas
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
            await semaphore.WaitAsync();
            try
            {
                await Task.Run(() =>
                {
                    var excelApp = new Application();
                    Workbook workbook = null;
                    Worksheet worksheet = null;

                    try
                    {
                        FileInfo fileInfo = new FileInfo(filePath);
                        fileInfo.IsReadOnly = false;

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
                                vcell.Value = komutText;
                            }

                            Range wcell = worksheet.Cells[i, "W"];
                            if (wcell.Value == null)
                            {
                                wcell.Value = ayrilmakomutText;
                            }
                        }

                        workbook.SaveAs(filePath);
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
                });
            }
            finally
            {
                semaphore.Release();
            }
        }

        public async Task<(double latitude, double longitude, double altitude)> GetLastGpsDataAsync(string filePath)
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

                            latitude = latCell.Value != null ? latCell.Value : 0;
                            longitude = lonCell.Value != null ? lonCell.Value : 0;
                            altitude = altCell.Value != null ? altCell.Value : 0;
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

                    return (latitude, longitude, altitude);
                });
            }
            finally
            {
                semaphore.Release();
            }
        }
    }
}
