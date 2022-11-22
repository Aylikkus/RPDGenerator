using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using RPDGenerator.Data;

namespace RPDGenerator.Tests
{
    [TestClass]
    public class FormatterTests
    {
        [TestMethod]
        public void GetSpecialization_TrimReturned()
        {
            string input = "  10.05.04  Информационно-аналитические системы безопасности  ";
            string expected = "10.05.04  Информационно-аналитические системы безопасности";

            string actual = Formatter.GetSpecialization(input);

            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void GetProfile_InQuoteReturn()
        {
            string input = "специализация N 1 \"Автоматизация информационно-аналитической деятельности\"";
            string expected = "Автоматизация информационно-аналитической деятельности";

            string actual = Formatter.GetProfile(input);

            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void GetEducationType_AfterSemicolonReturn()
        {
            string input = "Форма обучения: Очная";
            string expected = "Очная";

            string actual = Formatter.GetEducationType(input);

            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void GetEducationLevel_FormattedLevelReturn()
        {
            string inputSpec = "Квалификация: Специалист";
            string inputBac = "Квалификация: Бакалавр";
            string inputMas = "Квалификация: Магистрант";
            string inputAsp = "Квалификация: Аспирант";
            string inputUnknown = "Непонятная строка";

            string actualSpec = Formatter.GetEducationLevel(inputSpec);
            string actualBac = Formatter.GetEducationLevel(inputBac);
            string actualMas = Formatter.GetEducationLevel(inputMas);
            string actualAsp = Formatter.GetEducationLevel(inputAsp);
            string actualUnknown = Formatter.GetEducationLevel(inputUnknown);

            Assert.AreEqual("специалитет", actualSpec);
            Assert.AreEqual("бакалавриат", actualBac);
            Assert.AreEqual("магистратура", actualMas);
            Assert.AreEqual("аспирантура", actualAsp);
            Assert.AreEqual(null, actualUnknown);
        }
    }
}
