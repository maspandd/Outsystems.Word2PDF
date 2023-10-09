using NUnit.Framework;
using OutSystems.ExternalLibraries.SDK;
using System.Net.Mime;
using static System.Net.Mime.MediaTypeNames;

namespace UnitTest.WordPDF
{
    public class Tests
    {
        [Test]
        public void Test1()
        {
            var ConvertWordPDF = new WordToPDF();
            string filepath = Path.Combine("C:\\Users\\Lenovo\\source\\repos\\Word2PDF\\UnitTest.WordPDF\\resources\\test.docx"); //Add your path file here for testing
            byte[] bytes = File.ReadAllBytes(filepath);
            ConvertWordPDF.Doc2PDF(bytes, out byte[] PDF, out string Message, out int Code);

            byte[] filesave = PDF;
            string filePath = @"C:\Users\Lenovo\Documents\ResultO2.pdf";
            File.WriteAllBytes(filePath, filesave);

            Assert.True(Message == "");
        }
    }
}