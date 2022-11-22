using System;
using System.Text;
using System.Text.RegularExpressions;

namespace RPDGenerator.Data
{
    public static class Formatter
    {
        public static string GetSpecialization(string text)
        {
            return text.Trim();
        }

        public static string GetProfile(string text)
        {
            if (text.Contains("\""))
            {
                // Достаём текст из двойных кавычек
                var reg = new Regex("\".*?\"");
                var matches = reg.Matches(text);
                return matches[0].Value.Trim('\"');
            }
            else
                return text;
        }

        public static string GetEducationType(string text)
        {
            return text.Replace("Форма обучения: ", "");
        }

        public static string GetEducationLevel(string level)
        {
            switch (level.ToLowerInvariant().Replace("квалификация: ", ""))
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
}
