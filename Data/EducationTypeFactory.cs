using System;

namespace RPDGenerator.Data
{
    static class EducationTypeFactory
    {
        public static string GetType(string text)
        {
            return text.Replace("Форма обучения: ", "");
        }
    }
}
