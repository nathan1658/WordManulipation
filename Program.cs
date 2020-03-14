using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace WordManulipation
{
    class Program
    {
        static string labelPath = @"C:\Users\Nathan\source\repos\WordManulipation\WordManulipation\sample prescription\label.png";

        static void Main(string[] args)
        {
            var filePaths = Directory.GetFiles(@"C:\Users\Nathan\source\repos\WordManulipation\WordManulipation\sample prescription").Where(x => Path.GetExtension(x).Equals(".doc") || Path.GetExtension(x).Equals(".docx")).ToList();
            foreach(var filePath in filePaths)
            {
                Console.WriteLine("File: " + filePath);
                Console.ReadLine();
                doFile(filePath);
            }
         
        }


        static void doFile(string filePath)
        {
            var labelBytes = File.ReadAllBytes(labelPath);
            var labelStream = new MemoryStream();
            labelStream.Write(labelBytes, 0, labelBytes.Length);

            var wordStream = new MemoryStream();
            var fileExt = Path.GetExtension(filePath);
            if (fileExt.Equals(".doc"))
            {
                Aspose.Words.Document doc = new Aspose.Words.Document(filePath);
                doc.Save(wordStream, Aspose.Words.SaveFormat.Docx);
            }
            else if (fileExt.Equals(".docx"))
            {
                var fileByte = File.ReadAllBytes(filePath);
                wordStream.Write(fileByte, 0, fileByte.Length);
            }
            else
            {
                throw new Exception("Diu");
            }
            var resultStream = new MemoryStream();
            var productByte = WordManulipation.Helper.WordManipulation(wordStream.ToArray(), labelStream.ToArray());
            var fileName = $"{DateTime.Now.ToString("HHmmss")}.docx";
            File.WriteAllBytes(fileName, productByte);
            Process.Start(fileName);
        }
    }
}
