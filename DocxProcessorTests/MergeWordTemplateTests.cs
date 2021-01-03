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
            string TemplateFilePath = "C:\\Users\\lonye\\Desktop\\SideProject\\WordProcessor\\DocxProcessorTests\\WordTemplate\\test.docx";
            string TemplateFilePath2 = "C:\\Users\\lonye\\Desktop\\SideProject\\WordProcessor\\DocxProcessorTests\\WordTemplate\\test2.docx";
            string OutputFilePath = "C:\\Users\\lonye\\Desktop\\SideProject\\WordProcessor\\DocxProcessorTests\\WordTemplate\\MergeOutput.docx";
            FileStream docx1 = new FileStream(TemplateFilePath, FileMode.Open);
            FileStream docx2 = new FileStream(TemplateFilePath2, FileMode.Open);

            var Replacer = new ReplaceWordTemplate();
            var Merger = new MergeWordTemplate();

            List<Stream> Result = new List<Stream>();
            
            Dictionary<string, string> ReplaceItems = new Dictionary<string, string>();
            ReplaceItems.Add("#1#", "123");

            Result.Add(docx1);
                                    
            Result.Add(new MemoryStream(Replacer.Replace(OutputFilePath, ReplaceItems)));

            FileStream fs = new FileStream(OutputFilePath, FileMode.Create);

            BinaryWriter bw = new BinaryWriter(fs);

            bw.Write(Merger.MergeDocxsIntoOne(Result));

            bw.Close();

            fs.Close();
        }
    }
}

