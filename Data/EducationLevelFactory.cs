using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPDGenerator.Data
{
    class EducationLevelFactory
    {
        string _level;

        public EducationLevelFactory(string level)
        {
            _level = level.ToLowerInvariant().Replace("квалификация: ", "");
        }

        public string EducationLevel
        {
            get
            {
                switch (_level)
                {
                    case "бакалавр":
                        return "бакалавриат";
                    case "специалист":
                        return "специалитет";
                    case "магистрант":
                        return "магистратура";
                    case "аспирант":
                        return "аспирантура";
                    default:
                        return null;
                }
            }
        }

        public string GraduationLevel => _level;
    }
}
