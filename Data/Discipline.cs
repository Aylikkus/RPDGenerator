using System;

namespace RPDGenerator.Data
{
    struct Discipline
    {
        public string Code { get; set; }
        public string Name { get; set; }
        // public int ClockCount { get; set; }

        public Discipline(string code, string name)
        {
            Code = code;
            Name = name;
        }
    }
}
