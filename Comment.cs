using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace coursework
{
    internal class Comment
    {
        public string Text { get; set; }
        public bool ForGrade3 { get; set; }
        public bool ForGrade4 { get; set; }
        public bool ForGrade5 { get; set; }

        public Comment(string text, bool g3, bool g4, bool g5) {
            Text = text;
            ForGrade3 = g3;
            ForGrade4 = g4;
            ForGrade5 = g5;

        }
    }
}
