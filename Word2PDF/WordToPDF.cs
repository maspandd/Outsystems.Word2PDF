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


namespace Word2PDF
{
    public class WordToPDF : IWord2PDF
    {

        public void Doc2PDF_New(byte[] Doc, out byte[] PDF, out string Message, out int Code)
        {
            PDF = new byte[] { };
            Code = 0;
            Message = "";

            try
            {
                MemoryStream memoryStream = new MemoryStream();
                Stream stream = new MemoryStream(Doc);

                WordDocument document =  new WordDocument(stream, Syncfusion.DocIO.FormatType.Automatic);

                DocIORenderer render = new DocIORenderer();

                Syncfusion.Pdf.PdfDocument pdf = render.ConvertToPDF(document);

                render.Dispose();
                document.Dispose();

                var PDFBin = memoryStream.ToArray();

                PDF = PDFBin;
                Code = 200;
                Message = "OK"; 
                

            }
            catch(Exception ex) { Code = 500; Message = ex.Message; }
        }

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

                string result = System.Text.Encoding.UTF8.GetString(bytes);
                var test = result.Replace("Created with a trial version of Syncfusion Word library", " ");
                byte[] bitTest = System.Text.Encoding.UTF8.GetBytes(test);

                string result2 = System.Text.Encoding.UTF8.GetString(bitTest);
                var test2 = result2.Replace("or registered the wrong key in your", " ");
                byte[] bitTest2 = System.Text.Encoding.UTF8.GetBytes(test2);

                string result3 = System.Text.Encoding.UTF8.GetString(bitTest2);
                var test3 = result3.Replace(" application. Click </span><a style=\"color:#0000FF;font-family:Calibri;font-size:12pt;text-transform:none;font-weight:bold;font-style:normal;font-variant:normal;text-decoration: underline;\" href=\"https://www.syncfusion.com/account/claim-license-key?pl=ZmlsZWZvcm1hdHM=&amp;vs=MjMuMS4zNg==\"><span lang=\"en-US\" class=\"Default_Paragraph_Font Hyperlink\" style=\"color:#0000FF;font-family:Calibri;font-size:12pt;text-transform:none;font-weight:bold;font-style:normal;font-variant:normal;text-decoration: underline;\">here</span></a><span lang=\"en-US\" style=\"color:#FF0000;font-family:Calibri;font-size:12pt;text-transform:none;font-weight:bold;font-style:normal;font-variant:normal;text-decoration: none;\"> to obtain the valid key.", " ");
                byte[] bitTest3 = System.Text.Encoding.UTF8.GetBytes(test3);

                MemoryStream stream = new MemoryStream(bitTest);

                ConverterProperties converterProperties = new ConverterProperties();

                PdfWriter writer = new PdfWriter(stream);
                iText.Kernel.Pdf.PdfDocument pdf = new iText.Kernel.Pdf.PdfDocument(writer);
                pdf.SetDefaultPageSize(iText.Kernel.Geom.PageSize.DEFAULT);
                var document = HtmlConverter.ConvertToDocument(test3, pdf, converterProperties);
                document.Close();

                var pdfBin = stream.ToArray();

                PDF = pdfBin;



                Code = 200;
                Message = "Success";




            }
            catch (Exception ex)
            {
                Code = 500;
                Message = ex.Message;
            }
        }
    }
}