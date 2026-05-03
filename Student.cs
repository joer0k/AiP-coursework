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

        public string GetFIO {
            get
            {
                if (string.IsNullOrWhiteSpace(Name))
                    return "Студент";

                string[] parts = Name.Split(' ', StringSplitOptions.RemoveEmptyEntries);

                if (parts.Length >= 3)
                {
                    string lastName = parts[0];
                    string firstInitial = parts[1][0].ToString().ToUpper();
                    string surnameInitial = parts[2][0].ToString().ToUpper();

                    return $"{lastName}_{firstInitial}_{surnameInitial}";
                }

                if (parts.Length == 2)
                {
                    string lastName = parts[0];
                    string firstInitial = parts[1][0].ToString().ToUpper();

                    return $"{lastName}_{firstInitial}";
                }

                return parts[0];
            }
        }
    }
}
