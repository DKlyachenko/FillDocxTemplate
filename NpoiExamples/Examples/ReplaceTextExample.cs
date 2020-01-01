using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace NpoiExamples.Examples
{
    public class ReplaceTextExample
    {
        private Dictionary<string, string> PersonInfo { get; set; } = new Dictionary<string, string>();
        public void Run()
        {
            var templateFileName = @"templates\template.docx";
            string path = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), $"{templateFileName}");
            var currentTime = DateTime.Now;
            var outputFileName = String.Format("{0} {1} {2}.{3}.{4}.docx",
                "output", currentTime.ToShortDateString(), currentTime.Hour, currentTime.Minute, currentTime.Second);

            XWPFDocument doc;

            FillPersonInfo();

            try
            {
                using (FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    doc = new XWPFDocument(fileStream);
                    fileStream.Close();
                }

                ReplaceData(doc, PersonInfo);

                using (FileStream fileStreamNew = new FileStream(outputFileName, FileMode.CreateNew))
                {
                    doc.Write(fileStreamNew);
                    fileStreamNew.Close();
                }

                doc.Close();

                Console.WriteLine("Done");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private void FillPersonInfo()
        {
            PersonInfo["firstName"] = "John";
            PersonInfo["lastName"] = "Doe";
            PersonInfo["email"] = "johndoe@gmail.com";
        }

        private void ReplaceData(XWPFDocument doc, Dictionary<string, string> data)
        {
            foreach (XWPFParagraph p in doc.Paragraphs)
            {
                foreach (var dataItem in data.Keys)
                {
                    string oldText = p.Text;
                    string template = "{" + dataItem + "}";
                    string newText = oldText.Replace(template, data[dataItem]);
                    if (p.Text != null && p.Text.Contains(template))
                    {
                        p.ReplaceText(oldText, newText);
                    }
                }
            }
        }
    }
}
