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
        Worksheet _comps;
        Worksheet _plan;
        Range _range;

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

        Dictionary<string, string> parseCompetentionCell(string cellCnt, 
            in Dictionary<string, string> allComps)
        {
            string[] comps = cellCnt?.Split(new char[] {';', ' '}, StringSplitOptions.RemoveEmptyEntries);
            Dictionary<string, string> discComps = new Dictionary<string, string>();
            foreach(var c in comps)
            {
                discComps[c] = allComps[c];
            }

            return discComps;
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

            // Компетенции
            _comps = _workBook.Sheets[5];
            _range = _comps.UsedRange;
            object[,] compsArr = (object[,])_range.Value[XlRangeValueDataType.xlRangeValueDefault];

            Dictionary<string, string> allComps = new Dictionary<string, string>(128);
            for (int compRow = 1; compRow < compsArr.GetLength(0); compRow++)
            {
                if (compsArr[compRow, 2] != null)
                {
                    string compName = ((string)compsArr[compRow, 2]).Trim();
                    allComps[compName] = (string)compsArr[compRow, 4];
                }
            }

            // План
            _plan = _workBook.Sheets[4];
            _range = _plan.UsedRange;
            object[,] valArr = (object[,])_range.Value[XlRangeValueDataType.xlRangeValueDefault];

            List<Discipline> disciplines = new List<Discipline>(128);
            for (int discRow = 1; discRow < valArr.GetLength(0); discRow++)
            {
                string discCode = valArr[discRow, 2]?.ToString();
                string discName = valArr[discRow, 3]?.ToString();

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
                        string workInfoSems = (string)valArr[discRow, j];
                        if (workInfoSems == null) workInfoSems = "";

                        if (workInfoSems.Length > 0)
                            examInfos[j - 4] = new WorkInfo(si);

                        foreach (char c in workInfoSems)
                        {
                            int n = int.Parse(c.ToString(), NumberStyles.HexNumber);
                            examInfos[j - 4].SetOn(n, 0);
                        }
                    }

                    // Лекции, Лабы, Практики, Сам.Работы, Контроль
                    WorkInfo[] lesInfos = new WorkInfo[5];

                    for (int j = 18, s = 1; (string)valArr[3, j] == "з.е." && s <= 16; j += 7, s++)
                    {
                        parseLessonCell((string)valArr[discRow, j + 2], s, ref lesInfos[0], si);
                        parseLessonCell((string)valArr[discRow, j + 3], s, ref lesInfos[1], si);
                        parseLessonCell((string)valArr[discRow, j + 4], s, ref lesInfos[2], si);
                        parseLessonCell((string)valArr[discRow, j + 5], s, ref lesInfos[3], si);
                        parseLessonCell((string)valArr[discRow, j + 6], s, ref lesInfos[4], si);
                    }

                    Dictionary<string, string> comps = parseCompetentionCell(
                        (string)valArr[discRow, valArr.GetLength(1)], allComps);

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
                    disc.Control = lesInfos[4];
                    disc.Competentions = comps;
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
            da.Competentions = allComps;

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
            while (Marshal.ReleaseComObject(_comps) > 0) { }
            while (Marshal.ReleaseComObject(_plan) > 0) { }
            while (Marshal.ReleaseComObject(_range) > 0) { }

            _app = null;
            _workBook = null;
            _title = null;
            _comps = null;
            _plan = null;
            _range = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
