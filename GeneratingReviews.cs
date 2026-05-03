using System.Data;
using Xceed.Words.NET;
using Xceed.Document.NET;

namespace coursework
{
    internal class GeneratingReviews
    {
        private readonly GeneralData _generalData;
        private readonly List<Student> _students;
        private readonly List<Comment> _advantages;
        private readonly List<Comment> _disadvantages;
        public string signaturePath { get; }

        public string savePath { get; }


        public GeneratingReviews(GeneralData generalData, List<Student> students,
                             List<Comment> advantages, List<Comment> disadvantages,
                             string signaturePath, string savePath)
        {
            _generalData = generalData;
            _students = students;
            _advantages = advantages;
            _disadvantages = disadvantages;
            this.signaturePath = signaturePath;
            this.savePath = savePath;
        }
        //private void WriteDebugCommentsLog( //Функция для проверки правильности отображения комментариев, создает текстовый файл в котором указаны достоинства и недостатки в работе
        //    int grade,
        //    List<Comment> selectedAdvantages,
        //    List<Comment> selectedDisadvantages)
        //{
        //    string logPath = Path.Combine(savePath, "debug_comments.txt");

        //    string text =
        //        $"Оценка: {grade}{Environment.NewLine}" +
        //        $"Достоинства:{Environment.NewLine}" +
        //        string.Join(Environment.NewLine,
        //            selectedAdvantages.Select(x => $"{x.Text} | ColorKey = {x.ColorKey}")) +
        //        Environment.NewLine +
        //        $"Недостатки:{Environment.NewLine}" +
        //        string.Join(Environment.NewLine,
        //            selectedDisadvantages.Select(x => $"{x.Text} | ColorKey = {x.ColorKey}")) +
        //        Environment.NewLine +
        //        "----------------------------------------" +
        //        Environment.NewLine;

        //    File.AppendAllText(logPath, text);
        //}
        private (string advantagesText, string disadvantagesText) GetRandomComments(int grade)
        {
            var (advantagesCount, disadvantagesCount) = GetCommentCountsByGrade(grade);

            var availableAdvantages = _advantages
                .Where(comment => IsSuitableForGrade(comment, grade))
                .ToList();

            var availableDisadvantages = _disadvantages
                .Where(comment => IsSuitableForGrade(comment, grade))
                .ToList();

            if (availableAdvantages.Count < advantagesCount)
            {
                throw new Exception(
                    $"Недостаточно достоинств для оценки {grade}. " +
                    $"Нужно {advantagesCount}, найдено {availableAdvantages.Count}."
                );
            }

            if (availableDisadvantages.Count < disadvantagesCount)
            {
                throw new Exception(
                    $"Недостаточно недостатков для оценки {grade}. " +
                    $"Нужно {disadvantagesCount}, найдено {availableDisadvantages.Count}."
                );
            }

            const int maxAttempts = 1000;

            for (int attempt = 0; attempt < maxAttempts; attempt++)
            {
                var selectedAdvantages = availableAdvantages
                    .OrderBy(x => Guid.NewGuid())
                    .Take(advantagesCount)
                    .ToList();

                var selectedDisadvantages = availableDisadvantages
                    .OrderBy(x => Guid.NewGuid())
                    .Take(disadvantagesCount)
                    .ToList();

                bool hasConflict = HasColorConflict(selectedAdvantages, selectedDisadvantages);

                if (!hasConflict)
                {
                    //WriteDebugCommentsLog(grade, selectedAdvantages, selectedDisadvantages);

                    string advantagesText = string.Join(". ", selectedAdvantages.Select(x => x.Text));
                    string disadvantagesText = string.Join(". ", selectedDisadvantages.Select(x => x.Text));

                    return (advantagesText, disadvantagesText);
                }
            }

            throw new Exception($"Не удалось подобрать совместимые комментарии для оценки {grade}. Проверьте цвета конфликтующих достоинств и недостатков.");
        }

        private bool HasColorConflict(List<Comment> selectedAdvantages, List<Comment> selectedDisadvantages)
        {
            foreach (var advantage in selectedAdvantages)
            {
                foreach (var disadvantage in selectedDisadvantages)
                {
                    if (IsConflictColor(advantage.ColorKey) &&
                        advantage.ColorKey == disadvantage.ColorKey)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private bool IsConflictColor(int colorKey)
        {
            // 0 — нет заливки.
            // 16777215 — белый цвет.
            // Такие цвета не считаются конфликтующими.
            return colorKey != 0 && colorKey != 16777215;
        }

        private (int advantagesCount, int disadvantagesCount) GetCommentCountsByGrade(int grade)
        {
            return grade switch
            {
                5 => (4, 1),
                4 => (3, 2),
                3 => (2, 3),
                _ => throw new ArgumentException($"Для оценки {grade} не задано количество комментариев")
            };
        }

        private bool IsSuitableForGrade(Comment comment, int grade)
        {
            return grade switch
            {
                3 => comment.ForGrade3,
                4 => comment.ForGrade4,
                5 => comment.ForGrade5,
                _ => false
            };
        }

        private void ConvertToPDF(string docxPath) {
            string pdfPath = Path.ChangeExtension(docxPath, ".pdf");
            var wordApp = new Microsoft.Office.Interop.Word.Application();

            try
            {
                var document = wordApp.Documents.Open(docxPath);
                document.ExportAsFixedFormat(pdfPath, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
                document.Close(false);

            }
            finally {
                wordApp.Quit();
            }

        }

        public void CreateReviews() {

            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string originalTemplatePath = Path.Combine(baseDir, "revTemplate.docx");

            if (!File.Exists(originalTemplatePath))
            {
                throw new FileNotFoundException("Шаблон revTemplate.docx не найден рядом с программой!");
            }

            string tempTemplatePath = Path.Combine(Path.GetTempPath(), "template_temp.docx");

            try
            {

                File.Copy(originalTemplatePath, tempTemplatePath, true);

                foreach (var student in _students) {
                    string outputFileName = Path.Combine(savePath, $"Рецензия_{student.GetFIO}.docx");

                    var ReplaceMap = new Dictionary<string, string> {
                        {"{NAME}", student.Name },
                        {"{COURSE}", _generalData.course },
                        {"{DIRECTIONOFTRAINING}", _generalData.directionOfTraining },
                        {"{DIRECTIVITY}", _generalData.directivity },
                        {"{THEME}", student.TopicWork },
                        {"{ACADEMICTITLE}", _generalData.academicTitleAndPosition },
                        {"{TEACHER}", _generalData.Teacher },
                        {"{GRADE}", student.ConformityRating },
                        {"{DATE}", _generalData.Date },
                        {"{FORMOFEDUCATION}", _generalData.formOfEducation }
                    };
                    using (DocX doc = DocX.Load(tempTemplatePath)) {
                        foreach (var variable in ReplaceMap) { 
                            doc.ReplaceText(new StringReplaceTextOptions { 
                                SearchValue = variable.Key,
                                NewValue = variable.Value ?? ""
                            });
                        }

                        for (int i = 0; i < 8; i++) {
                            doc.ReplaceText(new StringReplaceTextOptions
                            {
                                SearchValue = "{GRADE" + $"{i + 1}" + "}",
                                NewValue = $"{student.Grade}"
                            });
                        }


                        var comments = GetRandomComments(student.Grade);

                        doc.ReplaceText(new StringReplaceTextOptions
                        {
                            SearchValue = "{ADVANTAGES}",
                            NewValue = comments.advantagesText
                        });

                        doc.ReplaceText(new StringReplaceTextOptions
                        {
                            SearchValue = "{DISADVANTAGES}",
                            NewValue = comments.disadvantagesText
                        });

                        if (File.Exists(signaturePath))
                        {
                            var image = doc.AddImage(signaturePath);

                            float height = 48.19f;
                            float width = 103.19f;

                            var picture = image.CreatePicture(height, width);

                            var bookmark = doc.Bookmarks["SIGNATURE"];
                            bookmark.Paragraph.AppendPicture(picture);
                            
                        }
                        doc.SaveAs(outputFileName);
                        ConvertToPDF(outputFileName);
                    }

                }
                
            }
            finally {
                if (File.Exists(tempTemplatePath)) File.Delete(tempTemplatePath);
            }

        }
    }
}
