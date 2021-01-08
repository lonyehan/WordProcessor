using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using System.IO;

namespace DocxProcessor.Tests
{
    [TestClass]
    public class InsertWordTemplateTests
    {
        [TestMethod]
        public void InsertTableTest()
        {            
            
            FileStream originFile = new FileStream("C:\\Users\\歐家豪\\source\\repos\\WordProcessor\\Template\\StudentList.docx", FileMode.Open);
            MemoryStream destination = new MemoryStream();

            originFile.CopyTo(destination);
            originFile.Close();

            using (var wordDoc = WordprocessingDocument.Open(destination, true))
            {
                Table Target = wordDoc.MainDocumentPart.Document.Body.Descendants<Table>().First(bms => bms.InnerText.Contains("#編號#") && bms.InnerText.Contains("#身分證統一編號#") && bms.InnerText.Contains("#姓名#"));
                TableRow TargetRow = wordDoc.MainDocumentPart.Document.Body.Descendants<TableRow>().First(bms => bms.InnerText.Contains("#編號#") && bms.InnerText.Contains("#身分證統一編號#") && bms.InnerText.Contains("#姓名#"));
                var Replacer = new ReplaceWordTemplate();
                var WordInserter = new InsertWordTemplate();

                // replace 
                Dictionary<string, string> keyValuePairs = new Dictionary<string, string>();
                keyValuePairs.Add("#編號#", "1234");
                

                WordInserter.InsertTableRow(Target, Replacer.Replace(TargetRow, keyValuePairs));

                wordDoc.Save();
            }

            destination.Position = 0;
                        

            string OutputFilePath = "C:\\Users\\歐家豪\\source\\repos\\WordProcessor\\Template\\testByModel.docx";
            FileStream fs = new FileStream(OutputFilePath, FileMode.Create);

            BinaryWriter bw = new BinaryWriter(fs);

            bw.Write(destination.ToArray());

            bw.Close();

            fs.Close();            
        }
    }
}
