using Microsoft.VisualStudio.TestTools.UnitTesting;
using RPDGenerator.Data;

namespace RPDGenerator.Tests
{
    [TestClass]
    public class SemesterInfoTests
    {
        SemesterInfo _semesterInfo;

        [TestInitialize]
        public void SemesterInfoTestsInitialize()
        {
            _semesterInfo = new SemesterInfo();
        }

        [TestMethod]
        public void SemestersAndCourses_AddSemester_MatchingReturn()
        {
            _semesterInfo.AddSemester(1);
            _semesterInfo.AddSemester(6);

            int[] expectedSem = new int[] { 1, 6 };
            int[] expectedCour = new int[] { 1, 3 };

            CollectionAssert.AreEqual(expectedSem, _semesterInfo.Semesters);
            CollectionAssert.AreEqual(expectedCour, _semesterInfo.Courses);
        }

        [TestMethod]
        public void SemestersAndCourses_AddCourses_MatchingReturn()
        {
            _semesterInfo.AddCourse(1);
            _semesterInfo.AddCourse(4);

            int[] expectedSem = new int[] { 1, 2, 7, 8 };
            int[] expectedCour = new int[] { 1, 4 };

            CollectionAssert.AreEqual(expectedSem, _semesterInfo.Semesters);
            CollectionAssert.AreEqual(expectedCour, _semesterInfo.Courses);
        }
    }
}
