using Microsoft.VisualStudio.TestTools.UnitTesting;
using RPDGenerator.Data;
using System;

namespace RPDGenerator.Tests
{
    [TestClass]
    public class DocAttributesTests
    {
        [TestMethod]
        public void ProfileAbbrevation_SAPRReturn()
        {
            string profile = "Системы автоматизированного проектирования";
            DocAttributes attributes = new DocAttributes("example", profile);

            string expected = "САПР";

            Assert.AreEqual(expected, attributes.ProfileAbbrevation);
        }

        [TestMethod]
        public void ProfileAbbrevation_AIDReturn()
        {
            string profile = "Автоматизация информационно-аналитической деятельности";
            DocAttributes attributes = new DocAttributes("example", profile);

            string expected = "АИД";

            Assert.AreEqual(expected, attributes.ProfileAbbrevation);
        }

        [TestMethod]
        public void ProfileAbbrevation_ISITVNPIOReturn()
        {
            string profile = "Информационные системы и технологии в науке, промышленности и образовании";
            DocAttributes attributes = new DocAttributes("example", profile);

            string expected = "ИСиТвНПиО";

            Assert.AreEqual(expected, attributes.ProfileAbbrevation);
        }
    }
}
