using System;
using System.Collections.Generic;
using System.Text;

namespace Simple.ToExcel.Classes
{
    public class Pirate
    {
        public Pirate(string name, string nickname, double bounty, bool captured, string level)
        {
            this.name = name;
            this.nickname = nickname;
            this.bounty = bounty;
            this.captured = captured;
            this.level = level;
        }

        public string name { get; set; }
        public string nickname { get; set; }
        public double bounty { get; set; }
        public bool captured { get; set; }
        public string level { get; set;}
    }
}
