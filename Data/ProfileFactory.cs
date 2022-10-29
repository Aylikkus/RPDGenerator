using System;

namespace RPDGenerator.Data
{
    static class ProfileFactory
    {
        public static string GetProfile(string text)
        {
            return text.Substring(text.IndexOf("\"") + 1, text.Length - 1);
        }
    }
}
