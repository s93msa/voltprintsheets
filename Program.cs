//using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
//using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;
//using ClosedXML.Excel;
//using PdfSharp.Drawing;
//using PdfSharp.Pdf;
//using PdfSharp.Pdf.IO;
//using Color = System.Drawing.Color;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Diagnostics;

namespace printSheets
{
    class Program
    {
        private string _path; // = @"C:\Temp\Jälla_2020-09-26\A\Lätt klass individuell grund och kür klass 15,16\";

        private const string pdfDir = "pdf";
        private const string mergedDir = "printPdf";

        //public string Path { get; set; }


        static void Main(string[] args)
        {

            //var fileNames = new List<string> {
            //    "1Hannah Hede_A_klass15_Grund_Ravnkjermindes Nanny_Sa",
            //    "2Hannah Hede_A_klass15_Kür_Ravnkjermindes Nanny_Sa" };

            var program = new Program();
            program.path(args);
            var fileNames = program.getFiles().ToArray();
            Array.Sort(fileNames, new CompareStrangAsNumber());
            program.createDirInPath(pdfDir);

            // Excel läser standardskrivaren när den startar. En långsam/offline nätverksskrivare
            // gör exporten mycket långsammare, så byt temporärt till "Microsoft Print to PDF".
            string originalPrinter = new PrinterSettings().PrinterName;
            bool printerChanged = false;
            ConsoleCancelEventHandler restoreOnCancel = (s, e) => RestorePrinter(originalPrinter, printerChanged);
            Console.CancelKeyPress += restoreOnCancel;
            ICollection<string> pdfFiles;
            try
            {
                printerChanged = SetPrinter(PdfPrinterName, originalPrinter);
                pdfFiles = program.printResultsExcelHandler(fileNames);
            }
            finally
            {
                RestorePrinter(originalPrinter, printerChanged);
                Console.CancelKeyPress -= restoreOnCancel;
            }

            program.createDirInPath(mergedDir);

            Console.WriteLine("PAth: " + pdfFiles.FirstOrDefault());
            string fileName = program.CreateFileNameFromPath();
            program.MergeMultiplePDFIntoSinglePDF(pdfFiles, fileName);

        }

        private const string PdfPrinterName = "Microsoft Print to PDF";

        [DllImport("winspool.drv", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool SetDefaultPrinter(string name);

        // Returnerar true om standardskrivaren faktiskt byttes (och alltså ska återställas).
        private static bool SetPrinter(string printer, string currentPrinter)
        {
            if (string.Equals(printer, currentPrinter, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
            if (!PrinterSettings.InstalledPrinters.Cast<string>().Contains(printer, StringComparer.OrdinalIgnoreCase))
            {
                Console.WriteLine($"Skrivaren \"{printer}\" finns inte, behåller \"{currentPrinter}\".");
                return false;
            }
            if (!SetDefaultPrinter(printer))
            {
                Console.WriteLine($"Kunde inte byta standardskrivare till \"{printer}\" (fel {Marshal.GetLastWin32Error()}).");
                return false;
            }
            Console.WriteLine($"Standardskrivare tillfälligt ändrad: \"{currentPrinter}\" -> \"{printer}\"");
            return true;
        }

        private static void RestorePrinter(string originalPrinter, bool printerChanged)
        {
            if (!printerChanged || string.IsNullOrEmpty(originalPrinter))
            {
                return;
            }
            if (SetDefaultPrinter(originalPrinter))
            {
                Console.WriteLine($"Standardskrivare återställd till \"{originalPrinter}\"");
            }
            else
            {
                Console.WriteLine($"VARNING: Kunde inte återställa standardskrivaren till \"{originalPrinter}\" (fel {Marshal.GetLastWin32Error()}).");
            }
        }

        // pathen C:\episerver\voltige\VoltigeClosedXML\output\utskrift_2023-07-06\A\Svår klass 1, Seniorlag SM klass- Grund/ 
        // ger tillbaka 
        // A_Svår klass 1, Seniorlag SM klass- Grund
        private string CreateFileNameFromPath()
        { 
            var backslashes = _path.Split('\\');
            var newFileName = backslashes.LastOrDefault(s => !string.IsNullOrWhiteSpace(s));
            newFileName = backslashes.LastOrDefault(s => !string.IsNullOrWhiteSpace(s) && s != newFileName) + "_" + newFileName;
            createDir(mergedDir);
            return newFileName;
        }

        private void path(string[] args)
        {
            if (args.Length > 0)
            {
                _path = Path.GetFullPath(args[0]);
            }
            else
            {
               // _path = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
                _path = Directory.GetCurrentDirectory();
            }

            _path = _path + '\\';
        }

        private ICollection<string> getFiles()
        {
            return Directory.GetFiles(_path, "*.xlsx");
        }

        private void createDirInPath(string name)
        {
            DirectoryInfo di = Directory.CreateDirectory(_path + name);
        }
        private void createDir(string name)
        {
            DirectoryInfo di = Directory.CreateDirectory(name);
        }

        // Övre gräns för antal Excel-instanser (varje instans är en egen EXCEL.EXE-process).
        // Fler än så ger sällan någon vinst eftersom disk och Excel-uppstart blir flaskhals.
        private const int MaxExcelInstances = 8;
        // Ungefärlig minnesåtgång per Excel-instans, och hur mycket minne som lämnas åt övriga program.
        private const long MemoryPerExcelInstanceMb = 300;
        private const long MemoryReserveMb = 1024;

        [StructLayout(LayoutKind.Sequential)]
        private class MEMORYSTATUSEX
        {
            public uint dwLength = (uint)Marshal.SizeOf(typeof(MEMORYSTATUSEX));
            public uint dwMemoryLoad;
            public ulong ullTotalPhys;
            public ulong ullAvailPhys;
            public ulong ullTotalPageFile;
            public ulong ullAvailPageFile;
            public ulong ullTotalVirtual;
            public ulong ullAvailVirtual;
            public ulong ullAvailExtendedVirtual;
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool GlobalMemoryStatusEx([In, Out] MEMORYSTATUSEX buffer);

        private static int GetPhysicalCoreCount()
        {
            try
            {
                int cores = 0;
                using (var searcher = new System.Management.ManagementObjectSearcher("SELECT NumberOfCores FROM Win32_Processor"))
                {
                    foreach (var item in searcher.Get())
                    {
                        cores += Convert.ToInt32(item["NumberOfCores"]);
                    }
                }
                if (cores > 0) return cores;
            }
            catch (Exception)
            {
                // WMI otillgängligt, använd uppskattning nedan.
            }
            // Antag hyperthreading: logiska / 2.
            return Math.Max(1, Environment.ProcessorCount / 2);
        }

        private static long GetAvailableMemoryMb()
        {
            var status = new MEMORYSTATUSEX();
            return GlobalMemoryStatusEx(status) ? (long)(status.ullAvailPhys / (1024 * 1024)) : -1;
        }

        // Välj antal parallella Excel-instanser utifrån fysiska kärnor och ledigt minne.
        private static int CalculateWorkerCount(int fileCount)
        {
            int cores = GetPhysicalCoreCount();
            // Lämna en kärna åt operativsystemet och övriga program.
            int byCpu = Math.Max(1, cores - 1);

            long availableMb = GetAvailableMemoryMb();
            int byMemory = availableMb < 0
                ? byCpu
                : (int)Math.Max(1, (availableMb - MemoryReserveMb) / MemoryPerExcelInstanceMb);

            int workers = Math.Max(1, Math.Min(Math.Min(byCpu, byMemory), Math.Min(MaxExcelInstances, fileCount)));
            Console.WriteLine($"Prestanda: {cores} fysiska kärnor, {availableMb} MB ledigt minne " +
                              $"(CPU tillåter {byCpu}, minne tillåter {byMemory}, max {MaxExcelInstances}) -> {workers} Excel-instanser för {fileCount} filer");
            return workers;
        }

        private ICollection<string> printResultsExcelHandler(IList<string> fileNames)
        {
            var sw = Stopwatch.StartNew();
            var pdfFiles = new string[fileNames.Count];
            int workerCount = CalculateWorkerCount(fileNames.Count);
            int nextIndex = -1;
            var errors = new List<Exception>();

            // Excel Interop kräver STA-trådar. Varje tråd får en egen Excel-instans
            // och plockar nästa fil från en gemensam räknare tills alla är klara.
            var threads = new List<System.Threading.Thread>();
            for (int w = 0; w < workerCount; w++)
            {
                var thread = new System.Threading.Thread(() =>
                {
                    try
                    {
                        ExportWorker(fileNames, pdfFiles, ref nextIndex);
                    }
                    catch (Exception e)
                    {
                        lock (errors) { errors.Add(e); }
                    }
                });
                thread.SetApartmentState(System.Threading.ApartmentState.STA);
                thread.Start();
                threads.Add(thread);
            }
            foreach (var thread in threads)
            {
                thread.Join();
            }

            if (errors.Count > 0)
            {
                throw new AggregateException("Export till PDF misslyckades", errors);
            }

            Console.WriteLine($"Exporterade {fileNames.Count} filer med {workerCount} Excel-instanser på {sw.Elapsed.TotalSeconds:F1} s");
            return pdfFiles;
        }

        private static void ExportWorker(IList<string> fileNames, string[] pdfFiles, ref int nextIndex)
        {
            Excel.Application app = null;
            Excel.Workbooks workbooks = null;
            try
            {
                app = new Excel.Application
                {
                    Visible = false,
                    ScreenUpdating = false,
                    DisplayAlerts = false,
                    EnableEvents = false,
                };
                workbooks = app.Workbooks;

                int i;
                while ((i = System.Threading.Interlocked.Increment(ref nextIndex)) < fileNames.Count)
                {
                    var fileName = fileNames[i];
                    var pdfFullPath = Path.Combine(
                        Path.GetDirectoryName(fileName),
                        pdfDir,
                        Path.GetFileNameWithoutExtension(fileName) + ".pdf");

                    Excel.Workbook book = null;
                    try
                    {
                        // UpdateLinks: 0 = uppdatera inte externa länkar vid öppning.
                        book = workbooks.Open(fileName, UpdateLinks: 0, ReadOnly: true, AddToMru: false);
                        book.ExportAsFixedFormat(
                            Excel.XlFixedFormatType.xlTypePDF,
                            pdfFullPath,
                            IncludeDocProperties: false,
                            OpenAfterPublish: false);
                        book.Close(SaveChanges: false);
                    }
                    finally
                    {
                        if (book != null) Marshal.FinalReleaseComObject(book);
                    }

                    pdfFiles[i] = pdfFullPath;
                    Console.WriteLine($"  [{i + 1}/{fileNames.Count}] {Path.GetFileName(pdfFullPath)}");
                }
            }
            finally
            {
                if (workbooks != null) Marshal.FinalReleaseComObject(workbooks);
                if (app != null)
                {
                    app.Quit();
                    Marshal.FinalReleaseComObject(app);
                }
            }
        }

        private void MergeMultiplePDFIntoSinglePDF(ICollection<string> pdfFiles, string fileName)
        {
            Console.WriteLine("PAth2: " + pdfFiles.FirstOrDefault());
            var now = DateTime.Now;
            PdfDocument document = new PdfDocument();

            foreach (string pdfFile in pdfFiles)
            {
                Console.WriteLine("PAth3: " + pdfFile);
                PdfDocument inputPDFDocument = PdfReader.Open(pdfFile, PdfDocumentOpenMode.Import);
                document.Version = inputPDFDocument.Version;

                foreach (PdfPage page in inputPDFDocument.Pages)
                {
                    document.AddPage(page);
                }
                // When document is add in pdf document remove file from folder  
                //System.IO.File.Delete(pdfFile);
            }
            // Set font for paging  
            //XFont font = new XFont("Verdana", 9);
            //XBrush brush = XBrushes.Black;
            //// Create variable that store page count  
            //string noPages = document.Pages.Count.ToString();

            document.Options.CompressContentStreams = true;
            document.Options.NoCompression = false;
            // In the final stage, all documents are merged and save in your output file path.  
            //var firstFile = pdfFiles.FirstOrDefault();
            //if (firstFile != null)
            //{
            //    var pdfFile = Path.GetFileNameWithoutExtension(firstFile);
            //    var printPdf = firstFile.Replace(pdfFile, "print");
            //var pdfFullPath = $"{_path}/{mergedDir}/print.pdf";
            var pdfFullPath = $"{mergedDir}/{fileName}.pdf";
            Console.WriteLine("PAth: " + pdfFullPath);
            document.Save(pdfFullPath);
              //);
            //}


        }
    }

}
