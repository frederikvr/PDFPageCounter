using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using CsvHelper;
using System.Globalization;

namespace PDFPageCounter
{
    class Program
    {
        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        static void Main(string[] args)
        {
            List<string> pdfs = new List<string>();
            //local
            //var directories = Directory.GetDirectories(@"E:\Customers\Fidea\PDFsTestDocuments");

            //on server fidea
            var directories = Directory.GetDirectories(@"D:\Docbyte\DocShifterTesting\PDFsTestDocuments");

            foreach (var directory in directories)
            {
                var listOfPdfsInDirectory = Directory.GetFiles(directory, "*.pdf");
                foreach (var pdf in listOfPdfsInDirectory)
                {
                    try
                    {
                        pdfs.Add(pdf);
                    }
                    catch (Exception ex)
                    {

                        log.Info("Exception " + ex + "caught");
                    }
                }
            }

            string timestamp = DateTime.UtcNow.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture);

            //output local
            //var outputDirectory = @"E:\Customers\Fidea\" + timestamp + "FideaPDFPagesCounted.csv";

            //output server fidea
            var outputDirectory = @"D:\Docbyte\DocShifterTesting\" + timestamp + "FideaPDFPagesCounted.csv";

            using (var writer = new StreamWriter(outputDirectory))
            using (var csv = new CsvWriter(writer))
                foreach (var pdf in pdfs)
                {
                    try
                    {
                        using (var document = PdfiumViewer.PdfDocument.Load(pdf))
                        {
                            string path = pdf.ToString();
                            int pos = path.LastIndexOf(@"\") + 1;
                            string nameOfPDF = path.Substring(pos, path.Length - pos);

                            csv.WriteRecord(new PDFRecord() { Name = nameOfPDF, AmountOfPages = document.PageCount });
                            csv.NextRecord();

                            log.Info(pdf.ToString() + " pageCount = " + document.PageCount);
                        }
                    }
                    catch (Exception ex)
                    {
                        log.Info("Exception " + ex + "caught");
                        csv.WriteRecord(new PDFRecord() { Name = ex.ToString(), AmountOfPages = 0 });
                        csv.NextRecord();
                    }
                }
        }
    }

    internal class PDFRecord
    {
        public string Name { get; set; }
        public int AmountOfPages { get; set; }
    }
}
