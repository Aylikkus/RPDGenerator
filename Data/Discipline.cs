using System;
using System.Text;

namespace RPDGenerator.Data
{
    public struct Discipline
    {
        string _abbr;

        void updateAbbrevation()
        {
            string[] words = Name.Split(new char[] { ' ', '-' },
                        StringSplitOptions.RemoveEmptyEntries);
            StringBuilder abbr = new StringBuilder();
            foreach (var w in words)
            {
                if (w.Length > 1)
                    abbr.Append(char.ToUpper(w[0]));
            }
            _abbr = abbr.ToString();
        }

        public string Code { get; }
        public string Name { get; }
        public string Abbrevation
        {
            get
            {
                if (_abbr == null)
                {
                    updateAbbrevation();
                }

                return _abbr;
            }
        }
        // public int ClockCount { get; set; }

        public Discipline(string code, string name) : this()
        {
            Code = code;
            Name = name;
        }
    }
}
