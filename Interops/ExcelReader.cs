using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

using RPDGenerator.Data;

namespace RPDGenerator.Interops
{
    public class ExcelReader : IDisposable
    {
        Application _app;
        Workbooks _workBooks;
        Workbook _workBook;
        Worksheet _title;
        Worksheet _plan;
        Range _planRange;

        /// <summary>
        /// Вытягивает информацию о проф. дисциплине из эксель файла,
        /// по определённому шаблону
        /// </summary>
        /// <param name="pathToFile">Путь к Эксель файлу/param>
        /// <returns></returns>
        public DocAttributes PullAttributes(string pathToFile)
        {
            _app = new Application();
            _workBooks = _app.Workbooks;
            _workBook = _workBooks.Open(pathToFile,
                Type.Missing, true);

            // Нумерация листов начинается с единицы
            // Титул
            _title = _workBook.Sheets[1];
            string departament    = ((Range)_title.Cells[27, 3]).Value2;
            string faculty        = ((Range)_title.Cells[28, 3]).Value2;
            string specialization = Formatter.GetSpecialization(
                ((Range)_title.Cells[19, 3]).Value2);
            string profile        = Formatter.GetProfile(
                ((Range)_title.Cells[20, 3]).Value2);

            string grLevel        = ((Range)_title.Cells[30, 2]).Value2;
            string edLevel        = Formatter.GetEducationLevel(grLevel);
            string edType         = Formatter.GetEducationType(
                ((Range)_title.Cells[32, 2]).Value2);
            int year              = int.Parse(((Range)_title.Cells[30, 21]).Value2);

            // План
            _plan = _workBook.Sheets[4];
            _planRange = _plan.UsedRange;
            object[,] valArr = (object[,])_planRange.Value[XlRangeValueDataType.xlRangeValueDefault];

            List<Discipline> disciplines = new List<Discipline>(128);
            for (int i = 1; i < valArr.GetLength(0); i++)
            {
                string code = valArr[i, 2]?.ToString();
                string name = valArr[i, 3]?.ToString();

                if (code is null || name is null)
                    continue;

                // Проверка на группировку дисциплин
                if (code[0] == 'Б' && !name.ToLowerInvariant().Contains("дисциплины"))
                {
                    Discipline disc = new Discipline(code, name);
                    disciplines.Add(disc);
                }
            }

            DocAttributes da = new DocAttributes();
            da.Departament = departament;
            da.Faculty = faculty;
            da.Specialization = specialization;
            da.Profile = profile;
            da.EducationLevel = edLevel;
            da.GraduationLevel = grLevel;
            da.EducationType = edType;
            da.YearOfEntrance = year;
            da.Disciplines = disciplines;

            return da;
        }

        public void Dispose()
        {
            _workBook.Close(false);
            _workBooks.Close();
            _app.Quit();

            // Ручное освобождение из-за COM-объектов
            while (Marshal.ReleaseComObject(_app) > 0) { }
            while (Marshal.ReleaseComObject(_workBook) > 0) { }
            while (Marshal.ReleaseComObject(_workBooks) > 0) { }
            while (Marshal.ReleaseComObject(_title) > 0) { }
            while (Marshal.ReleaseComObject(_plan) > 0) { }
            while (Marshal.ReleaseComObject(_planRange) > 0) { }

            _app = null;
            _workBook = null;
            _title = null;
            _plan = null;
            _planRange = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
