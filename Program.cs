using System.Diagnostics;
using ClosedXML.Excel;


namespace TimeTracker
{
    class Program
    {
        private static TimeSpan _stopwatch = TimeSpan.Zero;
        private static TimeSpan _stopwatchAway = TimeSpan.Zero;
        private static bool _isTracking = true; // Start tracking immediately
        private static readonly ManualResetEvent ExitEvent = new ManualResetEvent(false);

        private static void Main(string[] args)
        {
            Thread trackingThread = new Thread(TrackTime);
            trackingThread.Start();

            Console.WriteLine("Zeiterfassung gestartet. Drücken Sie Strg+C, um zu beenden.");

            Console.CancelKeyPress += (sender, e) => {
                e.Cancel = true;
                StopTracking();
                ExitEvent.Set();
            };
            ExitEvent.WaitOne();

            SaveToExcel();
            Console.WriteLine("Zeiterfassung beendet.");
        }

        private static void TrackTime()
        {
            StartTracking();
            while (!ExitEvent.WaitOne(0)) // Check for exit signal every second
            {
                if (_isTracking)
                {
                    if (Process.GetProcessesByName("LogonUI").Length > 0)
                    {
                        StopTracking(); // Stoppt die Zeitmessung, wenn der Sperrbildschirm angezeigt wird
                    }
                    else
                    {
                        Console.WriteLine($"Verstrichene Zeit: {_stopwatch}");
                        _stopwatch = _stopwatch.Add(new TimeSpan(0, 0, 1));
                        Thread.Sleep(1000);
                    }
                }
                else
                {
                    if (!(Process.GetProcessesByName("LogonUI").Length > 0))
                    {
                        Console.WriteLine("PC entsperrt. Soll die Zeit mitgezählt werden? (J/N)");
                        var answer = Console.ReadLine()?.ToUpper();

                        if (answer == "J")
                        {
                            _stopwatch += _stopwatchAway;
                        }
                        StartTracking();
                    }
                    Console.WriteLine($"Verstrichene Zeit Away: {_stopwatchAway}");
                    _stopwatchAway = _stopwatchAway.Add(new TimeSpan(0, 0, 1));
                    Thread.Sleep(1000);
                }
            }
        }

        private static void StartTracking()
        {
            _isTracking = true;
        }

        private static void StopTracking()
        {
            _isTracking = false;
        }

        private static void SaveToExcel()
        {
            const string excelFilePath = @"C:\Users\Jonas.Brueck\OneDrive - PlanB. GmbH\Desktop\test.xlsx";

            try
            {
                using (var workbook = new XLWorkbook(excelFilePath)) // Open or create workbook
                {
                    var worksheet = workbook.Worksheet(1); // Get the first worksheet
                    if (worksheet == null)
                    {
                        worksheet = workbook.Worksheets.Add("Zeiterfassung"); // Create if not exists
                    }

                    int lastRow;
                    // Find the last row with data
                    if (worksheet.LastRowUsed() is not null)
                    {
                        lastRow = worksheet.LastRowUsed()!.RowNumber() + 1;
                    }else

                    {
                        lastRow = 1;
                    }

                    // Write the data
                    worksheet.Cell(lastRow, 1).Value = DateTime.Today;
                    worksheet.Cell(lastRow, 2).Value = _stopwatch;

                    workbook.Save(); 
                }

                Console.WriteLine("Daten erfolgreich in Excel gespeichert.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fehler beim Speichern in Excel: {ex.Message}");
            }
        }
    }
}