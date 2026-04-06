using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using static System.Runtime.InteropServices.JavaScript.JSType;
namespace coursework
{
    internal class SystemInputData
    {
        public List<Student> _students = new List<Student>();
        public List<Comment> _advantages = new List<Comment>();
        public List<Comment> _disadvantages = new List<Comment>();
        public GeneralData _generalData;

        public void ParseExcelFile(string path)
        {
            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet();

                        foreach (DataTable table in result.Tables)
                        {
                            switch (table.TableName)
                            {
                                case "Общие данные":
                                    ReadGeneralData(table);
                                    break;

                                case "Недостатки":
                                    ReadComments(table, _disadvantages);
                                    break;

                                case "Достоинства":
                                    ReadComments(table, _advantages);
                                    break;

                                case "Список студентов":
                                    ReadStudent(table);
                                    break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка чтения Excel: " + ex.Message);
            }
        }


        public void ReadGeneralData(DataTable table)
        {

            string type = table.Rows[0][1]?.ToString();
            string course = table.Rows[3][1]?.ToString();
            string direction = table.Rows[4][1]?.ToString();
            string directivity = table.Rows[5][1]?.ToString();
            string formOfEducation = table.Rows[6][1]?.ToString();
            string teacherName = table.Rows[7][1]?.ToString();
            string academic = table.Rows[8][1]?.ToString();
            string group = table.Rows[9][1]?.ToString();


            string sDate = table.Rows[10][1]?.ToString();
            

            _generalData = new GeneralData(
                type,
                course,
                direction,
                directivity,
                academic,
                group,
                sDate,
                teacherName,
                formOfEducation
            );
        }

        public void ReadComments(DataTable table, List<Comment> list) {
            list.Clear();
            for (int i = 1; i < table.Rows.Count; i++) {
                var row = table.Rows[i];
                string text = row[0]?.ToString();
                if (!string.IsNullOrWhiteSpace(text)) {
                    
                    string formattedText = text.Trim();

                    if (formattedText.EndsWith("."))
                    {
                        formattedText = formattedText.TrimEnd('.').Trim();
                    }

                    if (!string.IsNullOrEmpty(formattedText) && char.IsLower(formattedText[0]))
                    {
                        formattedText = char.ToUpper(formattedText[0]) + formattedText.Substring(1);
                    }

                    list.Add(new Comment(
                                formattedText,
                                !string.IsNullOrWhiteSpace(row[1]?.ToString()), 
                                !string.IsNullOrWhiteSpace(row[2]?.ToString()),
                                !string.IsNullOrWhiteSpace(row[3]?.ToString()) 
                            ));
                }
            }
        }

        public void ReadStudent(DataTable table) {
            _students.Clear();
            for (int i = 1; i < table.Rows.Count; i++) {
                var row = table.Rows[i];
                string name = row[0]?.ToString() ?? "";

                if (string.IsNullOrEmpty(name)) continue;
                
                string topic = row[1]?.ToString() ?? "";

                if (int.TryParse(row[2]?.ToString() ?? "", out int grade)) {
                    _students.Add(new Student(name, topic, grade));
                }
            }
        }
    }
}
