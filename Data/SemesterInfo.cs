using System;
using System.Collections.Generic;

namespace RPDGenerator.Data
{
    public struct SemesterInfo
    {
        /// <summary>
        /// Представляет собой семестры, 
        /// на которых идёт дисциплина
        /// Например: 0b1001 -> на первом и четвёртом
        /// </summary>
        ushort _semesterFlag;

        void enableFlag(int number)
        {
            _semesterFlag |= (ushort)(1 << number - 1);
        }

        public void AddSemester(int number)
        {
            int size = sizeof(ushort);

            if (number < 1 && number > size * 8)
                return;

            enableFlag(number);
        }

        public void AddCourse(int number)
        {
            int size = sizeof(ushort) * 8 / 2;

            if (number < 1 && number > size)
                return;

            enableFlag(number * 2);
            enableFlag(number * 2 - 1);
        }

        public int[] Semesters
        {
            get
            {
                int size = sizeof(ushort) * 8;
                List<int> flagNums = new List<int>(size);

                for (int i = size - 1; i >= 0; i--)
                {
                    int bit = (_semesterFlag >> i) & 1;
                    if (bit == 1)
                        flagNums.Add(i + 1);
                }

                return flagNums.ToArray();
            }
        }

        public int[] Courses
        {
            get
            {
                int size = sizeof(ushort) * 8;
                List<int> flagNums = new List<int>(size);

                for (int i = size - 1; i >= 0; i -= 2)
                {
                    int bitOdd = (_semesterFlag >> i) & 1;
                    int bitEven = (_semesterFlag >> i - 1) & 1;
                    if (bitOdd == 1 || bitEven == 1)
                        flagNums.Add((i / 2) + (i % 2));
                }

                return flagNums.ToArray();
            }
        }
    }
}
