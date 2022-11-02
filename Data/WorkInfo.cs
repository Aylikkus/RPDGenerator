using System;
using System.Collections.Generic;

namespace RPDGenerator.Data
{
    public class WorkInfo
    {
        /// <summary>
        /// Ссылка на информацию с семестрами
        /// </summary>
        SemesterInfo _semesterInfo;

        /// <summary>
        /// Отображает семестр и соотв. кол-во часов
        /// </summary>
        Dictionary<int, int> _workHours;

        public void SetOn(int semester, int hours)
        {
            if (semester > _semesterInfo.Size)
                return;

            _workHours[semester] = hours;
            _semesterInfo.AddSemester(semester);
        }

        public IEnumerable<KeyValuePair<int, int>> HourEnumerator
        {
            get
            {
                return _workHours;
            }
        }

        public WorkInfo(SemesterInfo semInfo)
        {
            _semesterInfo = semInfo;
            _workHours = new Dictionary<int, int>();
        }
    }
}
