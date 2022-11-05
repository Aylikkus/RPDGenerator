using System;
using System.Collections.Generic;

namespace RPDGenerator.Data
{
    public class SemesterInfo
    {
        /// <summary>
        /// Отображает размер флаговой переменной
        /// в битах
        /// </summary>
        const int _size = sizeof(ushort) * 8;

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
            if (number < 1 && number > _size)
                return;

            enableFlag(number);
        }

        public void AddCourse(int number)
        {
            if (number < 1 && number > _size / 2)
                return;

            enableFlag(number * 2);
            enableFlag(number * 2 - 1);
        }

        public int Size
        {
            get
            {
                return _size;
            }
        }

        public int[] Semesters
        {
            get
            {
                List<int> flagNums = new List<int>(_size);

                for (int i = _size - 1; i >= 0; i--)
                {
                    int bit = (_semesterFlag >> i) & 1;
                    if (bit == 1)
                        flagNums.Add(i + 1);
                }

                flagNums.Reverse();
                return flagNums.ToArray();
            }
        }

        public int[] Courses
        {
            get
            {
                List<int> flagNums = new List<int>(_size);

                for (int i = _size - 1; i >= 0; i -= 2)
                {
                    int bitOdd = (_semesterFlag >> i) & 1;
                    int bitEven = (_semesterFlag >> i - 1) & 1;
                    if (bitOdd == 1 || bitEven == 1)
                        flagNums.Add((i / 2) + (i % 2));
                }

                flagNums.Reverse();
                return flagNums.ToArray();
            }
        }
    }
}
