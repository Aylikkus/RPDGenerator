using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Text;

namespace RPDGenerator.Data
{
    public class Discipline
    {
        string _abbr;

        void updateAbbrevation()
        {
            string[] words = Name.Split(new char[] { ' ', '-', },
                        StringSplitOptions.RemoveEmptyEntries);
            StringBuilder abbr = new StringBuilder();
            foreach (var w in words)
            {
                if (w.Length > 1)
                {
                    abbr.Append(char.ToUpper(w[0]));
                    if (w[0] == '(')
                        abbr.Append(char.ToUpper(w[1]));

                    if (w[w.Length - 1] == ')')
                        abbr.Append(')');
                }
            }
            _abbr = abbr.ToString();
        }

        public string Code { get; }
        public string Name { get; }
        public string Abbrevation
        {
            get
            {
                if (_abbr == null)
                {
                    updateAbbrevation();
                }

                return _abbr;
            }
        }

        /// <summary>
        /// Представляет собой информацию,
        /// связанную с семестрами
        /// </summary>
        public SemesterInfo Semester;

        /// <summary>
        /// Лекции
        /// </summary>
        public WorkInfo Lectures { get; set; }

        /// <summary>
        /// Лабораторные
        /// </summary>
        public WorkInfo Laboratory { get; set; }

        /// <summary>
        /// Практики
        /// </summary>
        public WorkInfo Practice { get; set; }

        /// <summary>
        /// Самостоятельные работы
        /// </summary>
        public WorkInfo Independent { get; set; }

        /// <summary>
        /// Экзамены
        /// </summary>
        public WorkInfo Exam { get; set; }

        /// <summary>
        /// Зачёты
        /// </summary>
        public WorkInfo Credits { get; set; }

        /// <summary>
        /// Зачёты с оценкой
        /// </summary>
        public WorkInfo RatedCredits { get; set; }

        /// <summary>
        /// Курсовые работы
        /// </summary>
        public WorkInfo CourseWorks { get; set; }

        /// <summary>
        /// Курсовые проекты
        /// </summary>
        public WorkInfo CourseProjects { get; set; }

        /// <summary>
        /// Расчётно-графические работы
        /// </summary>
        public WorkInfo RGR { get; set; }

        public Discipline(string code, string name)
        {
            Code = code;
            Name = name;
        }
    }
}
