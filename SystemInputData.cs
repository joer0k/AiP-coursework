using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace coursework
{
    internal class SystemInputData
    {
        public List<Student> students = new List<Student>();
        public List<Comment> advantages = new List<Comment>();
        public List<Comment> disadvantages = new List<Comment>();
        public GeneralData generalData = null!;

        public bool ParseExcelFile(string path)
        {
            Excel.Application excelApp = null;
            Excel.Workbooks workbooks = null;
            Excel.Workbook workbook = null;
            Excel.Sheets sheets = null;

            try
            {
                if (!File.Exists(path))
                {
                    MessageBox.Show("Файл Excel не найден.");
                    return false;
                }

                excelApp = new Excel.Application();
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;

                workbooks = excelApp.Workbooks;
                workbook = workbooks.Open(path, ReadOnly: true);
                sheets = workbook.Worksheets;

                for (int i = 1; i <= sheets.Count; i++)
                {
                    Excel.Worksheet sheet = null;

                    try
                    {
                        sheet = (Excel.Worksheet)sheets[i];

                        switch (sheet.Name)
                        {
                            case "Общие данные":
                                ReadGeneralData(sheet);
                                break;

                            case "Недостатки":
                                ReadComments(sheet, disadvantages);
                                break;

                            case "Достоинства":
                                ReadComments(sheet, advantages);
                                break;

                            case "Список студентов":
                                ReadStudent(sheet);
                                break;
                        }
                    }
                    finally
                    {
                        ReleaseComObject(sheet);
                    }
                }

                if (generalData == null)
                {
                    MessageBox.Show("Не найден лист 'Общие данные'.");
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка чтения Excel: " + ex.Message);
                return false;
            }
            finally
            {
                if (workbook != null)
                    workbook.Close(false);

                ReleaseComObject(sheets);
                ReleaseComObject(workbook);
                ReleaseComObject(workbooks);

                if (excelApp != null)
                    excelApp.Quit();

                ReleaseComObject(excelApp);

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public void ReadGeneralData(Excel.Worksheet sheet)
        {
            string type = GetCellText(sheet, 1, 2);
            string course = GetCellText(sheet, 4, 2);
            string direction = GetCellText(sheet, 5, 2);
            string directivity = GetCellText(sheet, 6, 2);
            string formOfEducation = GetCellText(sheet, 7, 2);
            string teacherName = GetCellText(sheet, 8, 2);
            string academic = GetCellText(sheet, 9, 2);
            string group = GetCellText(sheet, 10, 2);
            string sDate = GetCellText(sheet, 11, 2);

            generalData = new GeneralData(
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

        public void ReadComments(Excel.Worksheet sheet, List<Comment> list)
        {
            list.Clear();

            int lastRow = GetLastUsedRow(sheet);

            for (int i = 2; i <= lastRow; i++)
            {
                string text = GetCellText(sheet, i, 1);

                if (!string.IsNullOrWhiteSpace(text))
                {
                    string formattedText = text.Trim();

                    if (formattedText.EndsWith("."))
                    {
                        formattedText = formattedText.TrimEnd('.').Trim();
                    }

                    if (!string.IsNullOrEmpty(formattedText) && char.IsLower(formattedText[0]))
                    {
                        formattedText = char.ToUpper(formattedText[0]) + formattedText.Substring(1);
                    }

                    bool g3 = !string.IsNullOrWhiteSpace(GetCellText(sheet, i, 2));
                    bool g4 = !string.IsNullOrWhiteSpace(GetCellText(sheet, i, 3));
                    bool g5 = !string.IsNullOrWhiteSpace(GetCellText(sheet, i, 4));

                    int colorKey = GetCellColorKey(sheet, i, 1);

                    list.Add(new Comment(
                        formattedText,
                        g3,
                        g4,
                        g5,
                        colorKey
                    ));
                }
            }
        }

        public void ReadStudent(Excel.Worksheet sheet)
        {
            students.Clear();

            int lastRow = GetLastUsedRow(sheet);

            for (int i = 2; i <= lastRow; i++)
            {
                string name = GetCellText(sheet, i, 1);

                if (string.IsNullOrWhiteSpace(name))
                    continue;

                string topic = GetCellText(sheet, i, 2);
                string gradeText = GetCellText(sheet, i, 3);

                if (int.TryParse(gradeText, out int grade))
                {
                    students.Add(new Student(name, topic, grade));
                }
            }
        }

        private string GetCellText(Excel.Worksheet sheet, int row, int column)
        {
            Excel.Range cell = null;

            try
            {
                cell = (Excel.Range)sheet.Cells[row, column];
                return cell.Text?.ToString() ?? "";
            }
            finally
            {
                ReleaseComObject(cell);
            }
        }

        private int GetCellColorKey(Excel.Worksheet sheet, int row, int column)
        {
            Excel.Range cell = null;
            Excel.Interior interior = null;

            try
            {
                cell = (Excel.Range)sheet.Cells[row, column];
                interior = cell.Interior;

                int colorIndex = Convert.ToInt32(interior.ColorIndex);

                // -4142 означает отсутствие заливки
                if (colorIndex == -4142)
                    return 0;

                int color = Convert.ToInt32(interior.Color);

                // 16777215 — белый цвет, его тоже не считаем конфликтным
                if (color == 16777215)
                    return 0;

                return color;
            }
            finally
            {
                ReleaseComObject(interior);
                ReleaseComObject(cell);
            }
        }

        private int GetLastUsedRow(Excel.Worksheet sheet)
        {
            Excel.Range usedRange = null;
            Excel.Range rows = null;

            try
            {
                usedRange = sheet.UsedRange;
                rows = usedRange.Rows;

                return usedRange.Row + rows.Count - 1;
            }
            finally
            {
                ReleaseComObject(rows);
                ReleaseComObject(usedRange);
            }
        }

        private void ReleaseComObject(object obj)
        {
            if (obj != null)
            {
                try
                {
                    Marshal.ReleaseComObject(obj);
                }
                catch
                {
                    // Если объект уже освобождён, ничего не делаем
                }
            }
        }
    }
}