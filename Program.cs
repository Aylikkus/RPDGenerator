using RPDGenerator.Data;
using RPDGenerator.WordGenerator;
using System;

namespace RPDGenerator
{
    class Program
    {
        static void Main()
        {
            string wordpattern = "C:\\Users\\HONOR\\Desktop\\ТЗ\\Макет.docx";
            string excel = "C:\\Users\\HONOR\\Desktop\\ТЗ\\Excel\\2018\\очная\\10.05.04_ИАСБ_АИАД_УП(plx)_5.6_2018_~.xlsx";
            DocAttributes dc = ExcelReader.ExcelReader.PullAttributes(excel);
            WordGenerator.WordGenerator.WordHelper helper = new WordGenerator.WordGenerator.WordHelper(wordpattern);
            helper.Process(dc);
            Console.WriteLine("end");
            Console.ReadLine();
        }
    }
}
