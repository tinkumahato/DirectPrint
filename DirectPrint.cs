using Microsoft.Reporting.WinForms;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;

namespace BS_Office_Invoice.App_Code
{
    public static class DirectPrint
    {
        private static int m_currentPageIndex;
        private static IList<Stream> m_streams;

        public static Stream CreateStream(string name, string fileNameExtension, Encoding encoding, string mimeType, bool willSeek)
        {
            Stream stream = new MemoryStream();
            m_streams.Add(stream);
            return stream;
        }


        public static void Export(ReportViewer report, bool print = true)
        {
            double w = report.LocalReport.GetDefaultPageSettings().PaperSize.Width;
            double h = report.LocalReport.GetDefaultPageSettings().PaperSize.Height;
            string deviceInfo = "<DeviceInfo><OutputFormat>EMF</OutputFormat><PageWidth>" + (w / 96.0) + "in</PageWidth><PageHeight>" + (h / 96.0) + "in</PageHeight><MarginTop>0.0in</MarginTop><MarginLeft>0.0in</MarginLeft><MarginRight>0.0in</MarginRight><MarginBottom>0.0in</MarginBottom></DeviceInfo>";
            Warning[] warnings;
            m_streams = new List<Stream>();
            report.LocalReport.Render("Image", deviceInfo, CreateStream, out warnings);
            foreach (Stream stream in m_streams)
            {
                stream.Position = 0;
            }
            if (print)
            {
                Print(report);
            }
        }
        // Handler for PrintPageEvents
        public static void PrintPage(object sender, PrintPageEventArgs ev)
        {
            Metafile pageImage = new Metafile(m_streams[m_currentPageIndex]);
            Rectangle adjustedRect = new Rectangle(0, 0, ev.PageSettings.PaperSize.Width, ev.PageSettings.PaperSize.Height);
            ev.Graphics.FillRectangle(Brushes.White, adjustedRect);
            ev.Graphics.DrawImage(pageImage, adjustedRect);
            m_currentPageIndex++;
            ev.HasMorePages = (m_currentPageIndex < m_streams.Count);
        }

        public static void Print(ReportViewer report)
        {
            if (m_streams == null || m_streams.Count == 0)
                throw new Exception("Error: no stream to print.");
            PrintDocument printDoc = new PrintDocument();
            printDoc.PrinterSettings.PrinterName = report.PrinterSettings.PrinterName;
            printDoc.DefaultPageSettings.PaperSize = (System.Drawing.Printing.PaperSize)report.PrinterSettings.DefaultPageSettings.PaperSize;
            if (!printDoc.PrinterSettings.IsValid)
            {
                throw new Exception("Error: cannot find the default printer.");
            }
            else
            {
                printDoc.PrintPage += new PrintPageEventHandler(PrintPage);
                m_currentPageIndex = 0;
                printDoc.Print();
            }
        }

        public static void PrintToPrinter(this ReportViewer report)
        {
            Export(report);
        }

        public static void DisposePrint()
        {
            if (m_streams != null)
            {
                foreach (Stream stream in m_streams)
                    stream.Close();
                m_streams = null;
            }
        }
    }

}
