using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Text;

namespace RPDGenerator.Data
{
    public struct Discipline
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
        /// связанную с семестрами (на каких
        /// семестрах данная дисциплина, лекции,
        /// и соответствующие часы на каждый семестр,
        /// зачёты, экзамены и т.д.)
        /// </summary>
        public SemesterInfo Semester;

        public Discipline(string code, string name) : this()
        {
            Code = code;
            Name = name;
        }
    }
}
