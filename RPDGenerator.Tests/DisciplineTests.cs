using Microsoft.VisualStudio.TestTools.UnitTesting;
using RPDGenerator.Data;
using System;

namespace RPDGenerator.Tests
{
    [TestClass]
    public class DisciplineTests
    {
        [TestMethod]
        public void Abbrevation_TLPRReturn()
        {
            string discName = "Технологии личностно-профессионального развития";
            Discipline disc = new Discipline("example", discName);

            string expected = "ТЛПР";

            Assert.AreEqual(expected, disc.Abbrevation);
        }

        [TestMethod]
        public void Abbrevation_EReturn()
        {
            string discName = "Экономика";
            Discipline disc = new Discipline("example", discName);

            string expected = "Э";

            Assert.AreEqual(expected, disc.Abbrevation);
        }

        [TestMethod]
        public void Abbrevation_ScopeReturn()
        {
            string discName = "История (история России, всеобщая история)";
            Discipline disc = new Discipline("example", discName);

            string expected = "И(ИРВИ)";

            Assert.AreEqual(expected, disc.Abbrevation);
        }
    }
}
