using System;
using System.Collections.Generic;

namespace RPDGenerator.Data
{
    struct DocAttributes
    {
        /// <summary>
        /// Кафедра
        /// </summary>
        public string Departament { get; set; }
        /// <summary>
        /// Факультет
        /// </summary>
        public string Faculty { get; set; }
        /// <summary>
        /// Дисциплина
        /// </summary>
        public List<Discipline> Disciplines { get; set; }
        /// <summary>
        /// Специализация
        /// </summary>
        public string Specialization { get; set; }
        /// <summary>
        /// Направленность
        /// </summary>
        public string Profile { get; set; }
        /// <summary>
        /// Уровень образования
        /// </summary>
        public string EducationLevel { get; set; }
        /// <summary>
        /// Квалификация, присваиваемая по специальности
        /// </summary>
        public string GraduationLevel { get; set; }
        /// <summary>
        /// Форма обучения (очная, заочная)
        /// </summary>
        public string EducationType { get; set; }
        /// <summary>
        /// Год набора
        /// </summary>
        public int YearOfEntrance { get; set; }
    }
}
