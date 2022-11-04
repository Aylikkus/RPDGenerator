using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Globalization;
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

        void parseLessonCell(string cellCnt, int sem, ref WorkInfo les, in SemesterInfo si)
        {
            bool parsed = int.TryParse(cellCnt, out int hours);

            if (parsed)
            {
                if (les == null)
                    les = new WorkInfo(si);

                les.SetOn(sem, hours);
            }
        }

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
                string discCode = valArr[i, 2]?.ToString();
                string discName = valArr[i, 3]?.ToString();

                if (discCode is null || discName is null)
                    continue;

                // Проверка на группировку дисциплин
                if (discCode[0] == 'Б' && !discName.ToLowerInvariant().Contains("дисциплины"))
                {
                    SemesterInfo si = new SemesterInfo();

                    // Экз, Зач, ЗачОц, КурПр, КурРаб, РГР
                    WorkInfo[] examInfos = new WorkInfo[6];

                    for (int j = 4; j < 10; j++)
                    {
                        string workInfoSems = (string)valArr[i, j];
                        if (workInfoSems == null) workInfoSems = "";

                        if (workInfoSems.Length > 0)
                            examInfos[j - 4] = new WorkInfo(si);

                        foreach (char c in workInfoSems)
                        {
                            int n = int.Parse(c.ToString(), NumberStyles.HexNumber);
                            examInfos[j - 4].SetOn(n, 0);
                        }
                    }

                    // Лекции, Лабы, Практики, Сам.Работы
                    WorkInfo[] lesInfos = new WorkInfo[4];

                    for (int j = 18, s = 1; (string)valArr[3, j] == "з.е." && s <= 16; j += 7, s++)
                    {
                        parseLessonCell((string)valArr[i, j + 2], s, ref lesInfos[0], si);
                        parseLessonCell((string)valArr[i, j + 3], s, ref lesInfos[1], si);
                        parseLessonCell((string)valArr[i, j + 4], s, ref lesInfos[2], si);
                        parseLessonCell((string)valArr[i, j + 5], s, ref lesInfos[3], si);
                    }

                    Discipline disc = new Discipline(discCode, discName);
                    disc.Semester = si;
                    disc.Exam = examInfos[0];
                    disc.Credits = examInfos[1];
                    disc.RatedCredits = examInfos[2];
                    disc.CourseProjects = examInfos[3];
                    disc.CourseWorks = examInfos[4];
                    disc.RGR = examInfos[5];
                    disc.Lectures = lesInfos[0];
                    disc.Laboratory = lesInfos[1];
                    disc.Practice = lesInfos[2];
                    disc.Independent = lesInfos[3];
                    disciplines.Add(disc);
                }
            }

            DocAttributes da = new DocAttributes(specialization, profile);
            da.Departament = departament;
            da.Faculty = faculty;
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
