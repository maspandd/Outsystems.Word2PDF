using System;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using System.Reflection.Metadata;
using System.IO;
using iText.Html2pdf;
using iText.Kernel.Pdf;
using iText.Bouncycastle;
using iText.Bouncycastleconnector;
using Outsystems.Word2PDF;

namespace Word2PDF
{
    public class WordToPDF : IWord2PDF
    {

        RemoveWatermark removeWatermark = new RemoveWatermark();

        public void Doc2PDF(byte[] Doc, out byte[] PDF, out string Message, out int Code)
        {
            PDF = new byte[] { };
            Code = 0;
            Message = "";


            try
            {
                MemoryStream StreamHmtl = new MemoryStream();
                Stream strDoc = new MemoryStream(Doc);
                using (WordDocument documentW = new WordDocument(strDoc, FormatType.Word2013))
                {
                    HTMLExport export = new HTMLExport();
                    export.SaveAsXhtml(documentW, StreamHmtl);
                }

                byte[] bytes = StreamHmtl.ToArray();

                removeWatermark.removeWatermark(bytes, out byte[] bitBytes, out string Text);

                MemoryStream stream = new MemoryStream(bitBytes);

                ConverterProperties converterProperties = new ConverterProperties();

                PdfWriter writer = new PdfWriter(stream, new WriterProperties());
                iText.Kernel.Pdf.PdfDocument pdf = new iText.Kernel.Pdf.PdfDocument(writer.SetCompressionLevel(CompressionConstants.BEST_COMPRESSION));
                pdf.SetDefaultPageSize(iText.Kernel.Geom.PageSize.DEFAULT);
                var document = HtmlConverter.ConvertToDocument(Text, pdf, converterProperties);
                document.Close();

                var pdfBin = stream.ToArray();

                PDF = pdfBin;



                Code = 200;
                Message = "";




            }
            catch (Exception ex)
            {
                Code = 500;
                Message = ex.Message;
            }
        }
    }
}