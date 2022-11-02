using System;

namespace RPDGenerator.Data
{
    public class Person
    {
        /// <summary>
        /// Фамилия
        /// </summary>
        public string Surname { get; set; }

        /// <summary>
        /// Имя
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Отчество
        /// </summary>
        public string Patronymic { get; set; }

        /// <summary>
        /// Ученая степень
        /// </summary>
        public string AcademicDegree { get; set; }

        /// <summary>
        /// Ученое звание
        /// </summary>
        public string AcademicTitle { get; set; }

        /// <summary>
        /// Должность в вузе
        /// </summary>
        public string JobTitle { get; set; }
    }
}
