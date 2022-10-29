using System;
using Microsoft.Office.Interop.Excel;
using RPDGenerator.Data;

using Excel = Microsoft.Office.Interop.Excel;

namespace RPDGenerator.ExcelReader
{
    static class ExcelReader
    {
        /// <summary>
        /// Вытягивает информацию о проф. дисциплине из эксель файла,
        /// по определённому шаблону
        /// </summary>
        /// <param name="pathToFile">Путь к Эксель файлу/param>
        /// <returns></returns>
        public static DocAttributes PullAttributes(string pathToFile)
        {
            Application app = new Application();
            Workbook workBook = app.Workbooks.Open(pathToFile,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);

            // Нумерация листов начинается с единицы
            // Титул
            Worksheet title = workBook.Sheets[1];
            string departament    = ((Range)title.Cells[27, 3]).Value2;
            string faculty        = ((Range)title.Cells[28, 3]).Value2;
            string specialization = ((Range)title.Cells[19, 3]).Value2;
            string profile        = ((Range)title.Cells[20, 3]).Value2;

            var elf = new EducationLevelFactory(((Range)title.Cells[30, 2]).Value2);
            string edLevel        = elf.EducationLevel;
            string grLevel        = elf.GraduationLevel;
            string edType         = EducationTypeFactory.GetType(
                ((Range)title.Cells[32, 2]).Value2);
            int year              = int.Parse(((Range)title.Cells[30, 21]).Value2);

            // ПланСвод
            Worksheet plan = workBook.Sheets[3];

            DocAttributes da = new DocAttributes();
            da.Departament = departament;
            da.Faculty = faculty;
            da.Specialization = specialization;
            da.Profile = profile;
            da.EducationLevel = edLevel;
            da.GraduationLevel = grLevel;
            da.EducationType = edType;
            da.YearOfEntrance = year;
            return da;
        }
    }
}
