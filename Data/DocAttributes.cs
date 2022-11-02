using System;
using System.Collections.Generic;
using System.Text;

namespace RPDGenerator.Data
{
    public struct DocAttributes
    {
        string _abbr;

        void updateAttributes()
        {
            switch(Profile)
            {
                case "Системы автоматизированного проектирования":
                    _abbr = "САПР";
                    break;
                case "Информационные технологии и программные комплексы":
                    _abbr = "ИТиПК";
                    break;
                case "Автоматизация информационно-аналитической деятельности":
                    _abbr = "АИД";
                    break;
                case "Проектное управление в инновационной сфере":
                    _abbr = "ПУвИС";
                    break;
                case "Информационные системы и технологии в дизайне":
                    _abbr = "ИСиТвД";
                    break;
                default:
                    string[] words = Profile.Split(new char[] { ' ', '-' },
                        StringSplitOptions.RemoveEmptyEntries);
                    StringBuilder abbr = new StringBuilder();
                    foreach(var w in words)
                    {
                        if (w.Length > 1)
                            abbr.Append(char.ToUpper(w[0]));
                        else
                            abbr.Append(char.ToLower(w[0]));
                    }
                    _abbr = abbr.ToString();
                    break;
            }
        }

        /// <summary>
        /// Кафедра
        /// </summary>
        public string Departament { get; set; }

        /// <summary>
        /// Факультет
        /// </summary>
        public string Faculty { get; set; }

        /// <summary>
        /// Дисциплины
        /// </summary>
        public List<Discipline> Disciplines { get; set; }

        /// <summary>
        /// Специализация
        /// </summary>
        public string Specialization { get; }

        /// <summary>
        /// Направленность
        /// </summary>
        public string Profile { get; }

        /// <summary>
        /// Аббревиатура направленности, например САПР или АИД
        /// </summary>
        public string ProfileAbbrevation 
        { 
            get
            {
                if (_abbr == null)
                    updateAttributes();

                return _abbr;
            }
        }

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

        public DocAttributes(string specialization, string profile) : this()
        {
            _abbr = null;

            Specialization = specialization;
            Profile = profile;
        }
    }
}
