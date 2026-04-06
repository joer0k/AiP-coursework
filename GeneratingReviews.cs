using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;
using System.IO;
using System.Reflection;
using Xceed.Document.NET;
using Microsoft.Office.Interop.Word;
using System.Collections.Specialized;

namespace coursework
{
    internal class GeneratingReviews
    {
        private readonly GeneralData _generalData;
        private readonly List<Student> _students;
        private readonly List<Comment> _advantages;
        private readonly List<Comment> _disadvantages;
        public string _signaturePath { get; }

        public string _savePath { get; }

        public GeneratingReviews(GeneralData generalData, List<Student> students,
                             List<Comment> advantages, List<Comment> disadvantages,
                             string signaturePath, string savePath)
        {
            _generalData = generalData;
            _students = students;
            _advantages = advantages;
            _disadvantages = disadvantages;
            _signaturePath = signaturePath;
            _savePath = savePath;
        }

        private string GetRandomComments(List<Comment> source, int grade, int count = 4) {
            return string.Join(". ", source
                .Where(c => (grade == 5 && c.ForGrade5) || (grade == 4 && c.ForGrade4))
                .OrderBy(x => Guid.NewGuid())
                .Take(count)
                .Select(c => c.Text));
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
            //var assembly = Assembly.GetExecutingAssembly();

            //string resourceName = "coursework.revTemplate.docx";

            //string tempTemplatePath = Path.Combine(Path.GetTempPath(), "template_temp.docx");

            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string originalTemplatePath = Path.Combine(baseDir, "revTemplate.docx");

            if (!File.Exists(originalTemplatePath))
            {
                throw new FileNotFoundException("Шаблон revTemplate.docx не найден рядом с программой!");
            }

            string tempTemplatePath = Path.Combine(Path.GetTempPath(), "template_temp.docx");

            try
            {
                //using (Stream stream = assembly.GetManifestResourceStream(resourceName)) {
                //    if (stream == null) throw new Exception("Возникли проблемы с необходимым для создания рецензии шаблоном");

                //    using (FileStream fileStream = new FileStream(tempTemplatePath, FileMode.Create)) {
                //        stream.CopyTo(fileStream);
                //    }
                //}

                File.Copy(originalTemplatePath, tempTemplatePath, true);

                foreach (var student in _students) {
                    string outputFileName = Path.Combine(_savePath, $"Рецензия_{student.GetFirstName}.docx");

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

                        //var advantagesText = string.Join(". ", _advantages
                        //.Where(c => (student.Grade == 5 && c.ForGrade5) || (student.Grade == 4 && c.ForGrade4)).Take(4)
                        //.Select(c => c.Text));

                        var advantagesText = GetRandomComments(_advantages, student.Grade);

                        doc.ReplaceText(new StringReplaceTextOptions
                        {
                            SearchValue = "{ADVANTAGES}",
                            NewValue = advantagesText
                        });

                        //                 var disadvantagesText = string.Join(". ", _disadvantages
                        //.Where(c => (student.Grade == 5 && c.ForGrade5) || (student.Grade == 4 && c.ForGrade4)).Take(4)
                        //.Select(c => c.Text));

                        var disadvantagesText = GetRandomComments(_disadvantages, student.Grade);
                        doc.ReplaceText(new StringReplaceTextOptions
                        {
                            SearchValue = "{DISADVANTAGES}",
                            NewValue = disadvantagesText
                        });

                        if (File.Exists(_signaturePath))
                        {
                            var image = doc.AddImage(_signaturePath);

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
