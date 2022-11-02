using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Text;

namespace RPDGenerator.Data
{
    public struct Discipline
    {
        string _abbr;

        /// <summary>
        /// Представляет собой семестры, 
        /// на которых идёт дисциплина
        /// Например: 0b1001 -> на первом и четвёртом
        /// </summary>
        ushort _semesterFlag;

        /// <summary>
        /// Представляет собой курсы,
        /// на которых идёт дисциплина
        /// Например: 0b0101 -> на первом и третьем
        /// </summary>
        byte _courseFlag;

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
        // public int ClockCount { get; set; }
        private int[] arrayFromFlag(ushort flag)
        {
            ushort size = sizeof(ushort) * 8;
            List<int> flagNums = new List<int>(size);

            for (int i = size - 1; i >= 0; i--)
            {
                int bit = (flag >> i) & 1;
                if (bit == 1)
                    flagNums.Add(i + 1);
            }

            return flagNums.ToArray();
        }

        private int enableFlag(int flag, int number)
        {
            return flag | (1 << number - 1);
        }

        public int[] Semesters
        {
            get
            {
                return arrayFromFlag(_semesterFlag);
            }
        }

        public int[] Courses
        {
            get
            {
                return arrayFromFlag(_courseFlag);
            }
        }

        public void AddSemester(int number)
        {
            int size = sizeof(ushort);

            if (number < 1 && number > size * 8)
                return;

            _semesterFlag = (ushort)enableFlag(_semesterFlag, number);
            _courseFlag   = (byte)enableFlag(_courseFlag, number / 2 + number % 2);
        }

        public void AddCourse(int number)
        {
            int size = sizeof(byte);

            if (number < 1 && number > size * 8)
                return;

            _courseFlag = (byte)enableFlag(_courseFlag, number);

            _semesterFlag = (ushort)enableFlag(_semesterFlag, number * 2);
            _semesterFlag = (ushort)enableFlag(_semesterFlag, number * 2 - number % 2);
        }

        public Discipline(string code, string name) : this()
        {
            Code = code;
            Name = name;
        }
    }
}
