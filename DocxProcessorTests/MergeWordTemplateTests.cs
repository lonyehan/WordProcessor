using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Drawing;
using System.IO;

namespace DocxProcessor.Tests
{
    [TestClass]
    public class MergeWordTemplateTests
    {
        [TestMethod]
        public void MergeTwoDocx()
        {
            string TemplateFilePath = "C:\\Users\\歐家豪\\source\\repos\\WordProcessor\\Template\\StudentList.docx";
            string TemplateFilePath2 = "C:\\Users\\歐家豪\\source\\repos\\WordProcessor\\Template\\StudentList2.docx";
            string TemplateFilePath3 = "C:\\Users\\歐家豪\\source\\repos\\WordProcessor\\Template\\ReplaceByModelList.docx";
            string TemplateFilePath4 = "C:\\Users\\歐家豪\\source\\repos\\WordProcessor\\Template\\StudentList4.docx";
            string OutputFilePath = "C:\\Users\\歐家豪\\source\\repos\\WordProcessor\\Template\\MergeOutput.docx";
            //FileStream docx1 = new FileStream(TemplateFilePath, FileMode.Open);
            FileStream docx2 = new FileStream(TemplateFilePath2, FileMode.Open);
            FileStream docx3 = new FileStream(TemplateFilePath3, FileMode.Open);
            FileStream docx4 = new FileStream(TemplateFilePath4, FileMode.Open);

            var Replacer = new ReplaceWordTemplate();
            var Merger = new MergeWordTemplate();

            List<Stream> Result = new List<Stream>();
            Dictionary<string, string> keyValuePairs = new Dictionary<string, string>();

            keyValuePairs.Add("#1#", "123");

            Result.Add(new MemoryStream(Replacer.Replace(TemplateFilePath, keyValuePairs)));
                                    
            Result.Add(docx2);
            Result.Add(docx3);
            Result.Add(docx4);

            FileStream fs = new FileStream(OutputFilePath, FileMode.Create);

            BinaryWriter bw = new BinaryWriter(fs);

            bw.Write(Merger.MergeDocxsIntoOne(Result));

            bw.Close();

            fs.Close();
        }
    }
}

