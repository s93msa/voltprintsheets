//using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Drawing;
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

            var i = new Program();
            i.path(args);
            var fileNames = i.getFiles().ToArray();
            Array.Sort(fileNames, new CompareStrangAsNumber());
            i.createDir(pdfDir);
            var pdfFiles = i.printResultsExcelHandler(fileNames);
            i.createDir(mergedDir);
            Console.WriteLine("PAth: " + pdfFiles.FirstOrDefault());
            i.MergeMultiplePDFIntoSinglePDF(pdfFiles);

        }

        private void path(string[] args)
        {
            if (args.Length > 0)
            {
                _path = args[0];
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

        private void createDir(string name)
        {
            DirectoryInfo di = Directory.CreateDirectory(_path + name);
        }


        private ICollection<string> printResultsExcelHandler(ICollection<string> fileNames)
        {
            Excel.Application MyApp = null;
            Excel.Workbook MyBook = null;
            Excel.Workbooks workbooks = null;
            Excel.Worksheet MySheet = null;
            //bool preliminiaryResults = checkBox1.Checked;
            //string fullpath = Path.Combine(printedresultsFolder, filename);
            //string pdfFullPath = fullpath + ".pdf";
            //String noresults = ConfigurationManager.AppSettings["noresults"];

            //List<String> noresultsList = noresults.Split(',').ToList();
            var pdfFiles = new List<string>();
            try
            {
                MyApp = new Excel.Application
                {
                    Visible = false,
                    ScreenUpdating = false

                    //DisplayAlerts = true
                };
                workbooks = MyApp.Workbooks;
                foreach (var fileName in fileNames)
                {
                    var path = System.IO.Path.GetDirectoryName(fileName);
                    var pdfFile = System.IO.Path.GetFileNameWithoutExtension(fileName);
                    var pdfFullPath = $"{path}/{pdfDir}/{pdfFile}.pdf";
                    MyBook = workbooks.Open(fileName, ReadOnly: true);
                    MyBook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, pdfFullPath);

                    MyApp.DisplayAlerts = false;
                    MyBook.Close();
                    pdfFiles.Add(pdfFullPath);
                }
                MyApp.DisplayAlerts = true;
                MyApp.Quit();
            }
            catch (Exception e)
            {
                throw e;
            //    this.UpdateMessageTextBox($"Save to PDF failed for {className} : {e.Message}");
            //    showMessageBox(e.Message);
            }
            finally
            {
                //Marshal.FinalReleaseComObject(MySheet);
                //Marshal.FinalReleaseComObject(MyBook);
                //Marshal.FinalReleaseComObject(workbooks);
                //Marshal.FinalReleaseComObject(MyApp);
                MySheet = null;
                MyBook = null;
                workbooks = null;
                MyApp = null;
            }

           

            return pdfFiles;
        }

        private void MergeMultiplePDFIntoSinglePDF(ICollection<string> pdfFiles)
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
                var pdfFullPath = $"{_path}/{mergedDir}/print.pdf";
            Console.WriteLine("PAth: " + pdfFullPath);
            document.Save(pdfFullPath);
              //);
            //}


        }
    }

}
