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
            try
            {
                Application app = new Application();
                Workbook workBook = app.Workbooks.Open(pathToFile,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);

                // Нумерация листов начинается с единицы
                Worksheet title = workBook.Sheets[1];
                string departament = ((Range)title.Cells[27, 3]).Value2;

                DocAttributes da = new DocAttributes();
                da.Departament = departament;
                return da;
            }
            catch(Exception)
            {
                return new DocAttributes();
            }
        }
    }
}
