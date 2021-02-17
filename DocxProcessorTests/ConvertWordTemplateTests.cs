using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using DocxProcessor;

namespace DocxProcessor.Tests
{
    [TestClass]
    public class ConvertWordTemplateTests
    {
        [TestMethod]
        public void Case1()
        {
            //convert Docx to HTML
            string OutputFilePath = "C:\\Users\\歐家豪\\source\\repos\\WordProcessor\\Template\\StudentList.html";
            string TargetDocxPath = "C:\\Users\\歐家豪\\source\\repos\\WordProcessor\\Template\\StudentList.docx";            

            var Converter = new ConvertWordTemplate();
            var Replacer = new ReplaceWordTemplate();


            FileStream fs = new FileStream(OutputFilePath, FileMode.Create);

            BinaryWriter bw = new BinaryWriter(fs);
            
            bw.Write(Converter.WordToHtml(TargetDocxPath));

            bw.Close();

            fs.Close();
        }
        [TestMethod]
        public async System.Threading.Tasks.Task Case2Async()
        {
            //convert Docx to HTML
            string OutputFilePath = "C:\\Users\\歐家豪\\source\\repos\\WordProcessor\\Template\\StudentList.html";
            string TargetPdfPath = "C:\\Users\\歐家豪\\source\\repos\\WordProcessor\\Template\\StudentList.pdf";

            var Converter = new ConvertWordTemplate();
            var Replacer = new ReplaceWordTemplate();


            FileStream fs = new FileStream(TargetPdfPath, FileMode.Create);

            BinaryWriter bw = new BinaryWriter(fs);

            Dictionary<string, string> keyValuePairs = new Dictionary<string, string>();

            keyValuePairs.Add(" ", "　");

            bw.Write(await Converter.HtmlToPdf(OutputFilePath));

            bw.Close();

            fs.Close();
        }
    }
}
