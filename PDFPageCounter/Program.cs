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
            string timestamp = DateTime.UtcNow.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture);

            //INPUT

            //local
            //var listOfPdfsInDirectory = Directory.GetFiles(@"E:\Customers\Fidea\PDFPageCounterTestDoc", "*.pdf", SearchOption.AllDirectories);

            //fidea test
            var listOfPdfsInDirectory = Directory.GetFiles(@"C:\CustomerTesting\Fidea\Exportfolders\folderToPageCount\", "*.pdf", SearchOption.AllDirectories);

            //OUTPUT
            //output local
            //var outputDirectory = @"E:\Customers\Fidea\" + timestamp + "FideaPDFPagesCounted.csv";

            //output QUA server fidea
            //var outputDirectory = @"D:\Docbyte\DocShifterTesting\" + timestamp + "FideaPDFPagesCounted.csv";

            //output TST server fidea
            var outputDirectory = @"C:\CustomerTesting\Fidea\Exportfolders\" + timestamp + "FideaPDFPagesCounted.csv";

            using (var writer = new StreamWriter(outputDirectory))
            using (var csv = new CsvWriter(writer))
                foreach (var pdf in listOfPdfsInDirectory)
                {
                    int pos = pdf.LastIndexOf(@"\") + 1;
                    string nameOfPDF = pdf.Substring(pos, pdf.Length - pos);

                    try
                    {
                        using (var document = PdfiumViewer.PdfDocument.Load(pdf))
                        {

                            csv.WriteRecord(new PDFRecord()
                            {
                                Name = nameOfPDF,
                                AmountOfPages = document.PageCount
                            });
                            csv.NextRecord();

                            log.Info(pdf + " pageCount = " + document.PageCount);
                        }
                    }
                    catch (Exception ex)
                    {
                        log.Info("Exception " + ex + " caught");
                        csv.WriteRecord(new PDFRecord() { Name = nameOfPDF, AmountOfPages = -1 });
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
