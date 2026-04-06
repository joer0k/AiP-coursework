using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace coursework
{
    internal class Student
    {
        public string Name { get; set; }
        public string TopicWork { get; set; }
        public int Grade { get; set; }

        public List<string> SelectedAdvantages { get; set; } = new List<string>();
        public List<string> SelectedDisadvantages { get; set; } = new List<string>();

        public Student(string name, string topic, int grade) {
            Name = name;
            TopicWork = topic;
            Grade = grade;
        }

        public string ConformityRating {
            get
            {
                string[] rating = { "неудовлетворительно", "удовлетворительно", "хорошо", "отлично" };
                return (Grade >= 2 && Grade <= 5) ? $"{Grade} ({rating[Grade - 2]})" : "вне диапазона";
            }
        }

        public string GetFirstName {
            get {
                return Name?.Split(' ', StringSplitOptions.RemoveEmptyEntries)?.FirstOrDefault();
            }
        }
    }
}
