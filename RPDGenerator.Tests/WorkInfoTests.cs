using Microsoft.VisualStudio.TestTools.UnitTesting;
using RPDGenerator.Data;
using System;

namespace RPDGenerator.Tests
{
    [TestClass]
    public class WorkInfoTests
    {
        SemesterInfo _semesterInfo;
        WorkInfo _workInfo;

        [TestInitialize]
        public void WorkInfoTestsInitialize()
        {
            _semesterInfo = new SemesterInfo();
            _workInfo = new WorkInfo(_semesterInfo);
        }

        [TestMethod]
        public void SetOn_1sem5hours_SemesterInfoSemestersReturn()
        {
            _workInfo.SetOn(1, 5);

            int[] expectedSem = new int[] { 1 };

            CollectionAssert.AreEqual(expectedSem, _semesterInfo.Semesters);
        }

        [TestMethod]
        public void SetOn_1sem5hours2sem6hours_Total11Return()
        {
            _workInfo.SetOn(1, 5);
            _workInfo.SetOn(2, 6);

            int expected = 11;

            Assert.AreEqual(expected, _workInfo.Total);
        }

        [TestMethod]
        public void SetOn_1sem5hours_HoursOnSemester1Return5()
        {
            _workInfo.SetOn(1, 5);

            int expected = 5;

            Assert.AreEqual(expected, _workInfo.HoursOnSemester(1));
        }

        [TestMethod]
        public void SetOn_NonExistentSemester_HoursOnSemester5Return0()
        {
            int expected = 0;

            Assert.AreEqual(expected, _workInfo.HoursOnSemester(5));
        }
    }
}
