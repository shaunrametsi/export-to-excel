using System;
using System.Collections.Generic;
using System.Text;

namespace Simple.ToExcel.Classes
{
    public class Marine
    {
        public Marine(string name, string nickname, string position, string post)
        {
            this.name = name;
            this.nickname = nickname;
            this.position = position;
            this.post = post;
        }

        public string name { get; set; }
        public string nickname { get; set; }
        public string position { get; set; }
        public string post { get; set; }
    }
}
