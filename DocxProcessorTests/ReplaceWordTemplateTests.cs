﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Drawing;
using System.IO;

namespace DocxProcessor.Tests
{
    [TestClass]
    public class ReplaceWordTemplateTests
    {
        [TestMethod]
        public void ReplaceCase1()
        {
            string TemplateFilePath = "C:\\Users\\歐家豪\\source\\repos\\WordProcessor\\Template\\StudentList.docx";
            string OutputFilePath = "C:\\Users\\歐家豪\\source\\repos\\WordProcessor\\Template\\StudentList2.docx";
            Dictionary<string, string> keyValuePairs = new Dictionary<string, string>();
            string TestStr = @"1.	符合報名條件及門檻者，依選校登記序號現場分發。
2.	本校升學績效優質，超高國立大學錄取率：109年第17屆畢業班，國立大學錄取率高達96%。
3.	課程以分組教學，並包含多種適性多元課程。擁有全新數位藝術與設計教室，設計與電繪課程、版畫課程、插畫創意風格課程與素描、水彩、水墨書畫等專業課程；設備、師資與課程規劃最健全，教學與輔導最用心!
4.	備有縝密的專車路線與4人一寢冷氣宿舍，優質環境歡迎蒞校參觀或來電詢問(037-868680分機204)。
";

            keyValuePairs.Add("#sbd_1_0_0#", TestStr);
            keyValuePairs.Add("#sbd_1_0_1#", "H12313213221G");
            var Replacer = new ReplaceWordTemplate();

            FileStream fs = new FileStream(OutputFilePath, FileMode.Create);

            BinaryWriter bw = new BinaryWriter(fs);

            bw.Write(Replacer.Replace(TemplateFilePath, keyValuePairs));

            bw.Close();

            fs.Close();
        }
        [TestMethod]
        public void ReplaceCase2()
        {
            string TemplateFilePath = "C:\\Users\\歐家豪\\source\\repos\\WordProcessor\\Template\\ApplyFromArt.docx";
            string OutputFilePath2 = "C:\\Users\\歐家豪\\source\\repos\\WordProcessor\\Template\\testCase2.docx";
            Dictionary<string, string> keyValuePairs = new Dictionary<string, string>();


            keyValuePairs.Add("#測試#", "替換");
            var Replacer = new ReplaceWordTemplate();

            Dictionary<string, string> keyValuePairs2 = new Dictionary<string, string>();

            keyValuePairs2.Add("#圖片#", "123");

            FileStream fs = new FileStream(OutputFilePath2, FileMode.Create);

            BinaryWriter bw = new BinaryWriter(fs);

            bw.Write(Replacer.Replace(Replacer.Replace(TemplateFilePath, keyValuePairs), keyValuePairs2));

            bw.Close();

            fs.Close();
        }
        [TestMethod]
        public void ReplaceByModel()
        {
            string TemplateFilePath = "C:\\Users\\歐家豪\\source\\repos\\WordProcessor\\Template\\StudentList.docx";
            string OutputFilePath = "C:\\Users\\歐家豪\\source\\repos\\WordProcessor\\Template\\StudentList2.docx";
            TestModel test = new TestModel();
            test.sbd_1_0_0 = "200";
            test.sbd_1_0_1 = "0905337291";            

            var Replacer = new ReplaceWordTemplate();
            FileStream fs = new FileStream(OutputFilePath, FileMode.Create);

            BinaryWriter bw = new BinaryWriter(fs);

            bw.Write(Replacer.Replace(TemplateFilePath, test));

            bw.Close();

            fs.Close();
        }
        [TestMethod]
        public void ReplaceByImage()
        {
            string TemplateFilePath = "C:\\Users\\歐家豪\\source\\repos\\WordProcessor\\Template\\AdmissionNumberForm.docx";

            string OutputFilePath = "C:\\Users\\歐家豪\\source\\repos\\WordProcessor\\Template\\testReplaceByImage.docx";

            var Replacer = new ReplaceWordTemplate();

            FileStream fs = new FileStream(OutputFilePath, FileMode.Create);

            BinaryWriter bw = new BinaryWriter(fs);

            Dictionary<string, ImageData> keyValuePairs = new Dictionary<string, ImageData>();

            ImageData Value = new ImageData("C:\\Users\\歐家豪\\source\\repos\\WordProcessor\\Template\\測試用證件照.jpg", Width: 4.1M, Height: 4.1M);

            keyValuePairs.Add("#證件照#", Value);

            bw.Write(Replacer.ReplaceTableCellByImage(TemplateFilePath, keyValuePairs));

            bw.Close();

            fs.Close();
        }
        [TestMethod]
        public void ReplaceByImage2()
        {
            string TemplateFilePath = "C:\\Users\\歐家豪\\source\\repos\\WordProcessor\\Template\\ApplyFromArt.docx";

            string OutputFilePath = "C:\\Users\\歐家豪\\source\\repos\\WordProcessor\\Template\\testReplaceByImage2.docx";

            var Replacer = new ReplaceWordTemplate();

            FileStream fs = new FileStream(OutputFilePath, FileMode.Create);

            BinaryWriter bw = new BinaryWriter(fs);

            Dictionary<string, ImageData> keyValuePairs = new Dictionary<string, ImageData>();

            string FileStorageRootPath = "D:\\Project\\files\\permission\\";
            string NewPath = Convert.ToString(Convert.ToInt64(3610), 16).PadLeft(8, '0');
            for (int i = 0; i < 10; i += 3)
            {
                NewPath = NewPath.Insert(i, "\\");
            }
            string FileNewPath = FileStorageRootPath + NewPath + ".jpg";

            ImageData Value = new ImageData(FileNewPath, Width: 4.1M, Height: 4.1M);

            keyValuePairs.Add("#圖片#", Value);

            bw.Write(Replacer.ReplaceTableCellByImage(TemplateFilePath, keyValuePairs));

            bw.Close();

            fs.Close();
        }
        [TestMethod]
        public void ReplaceByImage3()
        {
            string TemplateFilePath = "C:\\Users\\歐家豪\\Pictures\\藝才\\測試用證件照.jpg";

            Image img = Image.FromFile(TemplateFilePath);
            ImageData test = new ImageData(TemplateFilePath, 4.1M);
            //Console.Write(img.Width);
            //Console.Write((test.WidthInEMU / 360000M).ToString());
            Console.Write((test.HeightInEMU / 360000M).ToString());
            //1476000.0
            //81279349.86109869832640323690
        }
        [TestMethod]
        public void ReplaceByModelList()
        {
            string TemplateFilePath = "C:\\Users\\歐家豪\\source\\repos\\WordProcessor\\Template\\StudentList5.docx";
            string OutputFilePath = "C:\\Users\\歐家豪\\source\\repos\\WordProcessor\\Template\\ReplaceByModelList.docx";

            // ModelList
            List<Student> Datas = new List<Student>();
            Datas.Add(new Student());
            Datas.Add(new Student());
            Datas[0].sbd_1_0_0 = "1232";
            Datas[0].sbd_1_0_1  = "1232";
            Datas[0].sbd_1_0_2  = "1232";
            Datas[1].sbd_1_0_0 = "1232";
            Datas[1].sbd_1_0_1 = "1232";
            Datas[1].sbd_1_0_2 = "1232";
            
            var Replacer = new ReplaceWordTemplate();

            FileStream fs = new FileStream(OutputFilePath, FileMode.Create);

            BinaryWriter bw = new BinaryWriter(fs);


            bw.Write(Replacer.Replace(TemplateFilePath, Datas));            

            bw.Close();

            fs.Close();
        }
        [TestClass]
        public class TestModel
        {
            public string sbd_1_0_0 { get; set; }
            public string sbd_1_0_1 { get; set; }

        }
        [TestClass]
        public class Student
        {              
            public string sbd_1_0_0 { get; set; }
            public string sbd_1_0_1 { get; set; }
            public string sbd_1_0_2 { get; set; }
            //public string sbd_1_0_3 { get; set; }
            //public string sbd_1_0_4 { get; set; }
            //public string sbd_1_0_5 { get; set; }
            //public string sbd_1_0_6 { get; set; }
            //public string sbd_1_0_7 { get; set; }
            //public string sbd_1_0_8 { get; set; }
            //public string sbd_1_0_9 { get; set; }
            //public string sbd_1_0_10 { get; set; }
            //public string sbd_1_0_11 { get; set; }
            //public string sbd_1_0_12 { get; set; }
            //public string sbd_1_0_13 { get; set; }
            //public string sbd_1_0_14 { get; set; }
            //public string sbd_1_0_15 { get; set; }
        }
    }
}
