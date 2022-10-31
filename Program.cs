using System;
using System.IO;
using RPDGenerator.Data;
using RPDGenerator.Interops;

namespace RPDGenerator
{
    class Program
    {
        static void Main()
        {
            string curDir = Environment.CurrentDirectory;
            string prjName = "RPDGenerator";
            string projectRoot = curDir.Substring(0, curDir.IndexOf(prjName) + prjName.Length);

            string wordpattern = projectRoot + "\\Макет.docx";
            string excel = projectRoot + "\\Excel\\2022\\очная\\10.05.04_ИАСБ_аиад_С_5,6_2022_очная.p~.xlsx";

            DocAttributes dc;
            using (ExcelReader er = new ExcelReader())
                dc = er.PullAttributes(excel);
            using (WordGenerator helper = new WordGenerator())
                helper.GenerateDocs(dc, wordpattern);
            Console.WriteLine("end");
            Console.ReadLine();
        }
    }
}
