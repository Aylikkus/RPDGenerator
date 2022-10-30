using System;
using RPDGenerator.Data;
using RPDGenerator.Interops;

namespace RPDGenerator
{
    class Program
    {
        static void Main()
        {
            string wordpattern = "C:\\Users\\4l1kk\\source\\repos\\RPDGenerator\\Макет.docx";
            string excel = "C:\\Users\\4l1kk\\source\\repos\\RPDGenerator\\Excel\\2018\\очная\\10.05.04_ИАСБ_АИАД_УП(plx)_5.6_2018_~.xlsx";
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
