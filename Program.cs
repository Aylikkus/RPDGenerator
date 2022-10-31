using System;
using RPDGenerator.Data;
using RPDGenerator.Interops;

namespace RPDGenerator
{
    class Program
    {
        static void Main()
        {
            string wordpattern = "C:\\Users\\HONOR\\Desktop\\ТЗ\\Макет.docx";
            string excel = "C:\\Users\\HONOR\\Desktop\\ТЗ\\Excel\\2022\\очная\\10.05.04_ИАСБ_аиад_С_5,6_2022_очная.p~.xlsx";
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
