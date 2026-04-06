using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace coursework
{
    internal class GeneralData
    {
        public readonly string type;
        public readonly string course;
        public readonly string directionOfTraining;
        public readonly string directivity;
        private string _teacher;
        public readonly string academicTitleAndPosition;
        public readonly string group;
        private string _date;
        public readonly string formOfEducation;

        private string _lastNameTeacher;
        private string _firstNameTeacher;
        private string _surnameTeacher;
        private string _initialsAndFirstNameTeacher;


        public GeneralData(string type, string course, string direction, string directivity,
                       string academic, string group, string date, string teacherName, string formOfEducation)
        {
            this.type = type;
            this.course = course;
            this.directionOfTraining = direction;
            this.directivity = directivity;
            this.academicTitleAndPosition = academic;
            this.group = group;
            this.Date = date;
            this.formOfEducation = formOfEducation;
            this.Teacher = teacherName;
        }

        public string Teacher
        {
            get => _teacher;
            set {
                if (_teacher != null) return;
                _teacher = value;

                if (!string.IsNullOrWhiteSpace(_teacher))
                {
                    string[] parts = _teacher.Split(new[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length >= 3) {
                        _firstNameTeacher = parts[0];
                        _lastNameTeacher = parts[1];
                        _surnameTeacher = parts[2];

                        _initialsAndFirstNameTeacher = $"{_firstNameTeacher[0]}.{_surnameTeacher[0]}. {_lastNameTeacher[0]}";
                }
                }
            }
        }

        public string Date
        {
            get => _date;
            set
            {
                if (string.IsNullOrWhiteSpace(value)) return;
                if (DateTime.TryParse(value, out DateTime parseddate))
                {
                    _date = parseddate.ToString("dd.MM.yyyy");

                }
                else {

                    _date = value;
                }

            }
        }

    }

}
