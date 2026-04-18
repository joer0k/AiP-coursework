using System.Diagnostics;

namespace coursework
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            openTemplateDialog.Filter = "Excel files(.xlsx, .xls)|*.xlsx;*.xls";
            openTemplateDialog.Title = "Выберите файл с исходными данными";

            openSignatureDialog.Filter = "Все изображения|*.jpg;*.jpeg;*.png|" +
                         "JPEG файлы (*.jpg;*.jpeg)|*.jpg;*.jpeg|" +
                         "PNG файлы (*.png)|*.png";
            openSignatureDialog.Title = "Выберите картинку вашей подписи";

        }

        //private void SetUIEnabled(bool enabled)
        //{
        //    btnCreate.Enabled = enabled;
        //    btnTemplate.Enabled = enabled;
        //    btnSignature.Enabled = enabled;
        //    btnWay.Enabled = enabled;

        //    this.Cursor = enabled ? Cursors.Default : Cursors.WaitCursor;
        //}

        private void btnTemplate_Click(object sender, EventArgs e)
        {
            if (openTemplateDialog.ShowDialog() == DialogResult.Cancel){ return; }
            string filename = openTemplateDialog.FileName;

            string templatepath = System.IO.Path.GetFullPath(filename);
            tbTemplate.Text = templatepath;
        }


        private void btnSignature_Click(object sender, EventArgs e)
        {
            if (openSignatureDialog.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }
            string filename = openSignatureDialog.FileName;

            string signaturepath = System.IO.Path.GetFullPath(filename);
            tbSignature.Text = signaturepath;
        }


        private void btnWay_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog.ShowDialog() == DialogResult.Cancel) { return; }
            string pathtosave = folderBrowserDialog.SelectedPath;
            tbWay.Text = pathtosave;

        }
        private async void btnCreate_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(tbTemplate.Text) ||
                string.IsNullOrWhiteSpace(tbSignature.Text) ||
                string.IsNullOrWhiteSpace(tbWay.Text))
            {
                MessageBox.Show("Заполните все поля!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }



            try
            {
                string templatepath = tbTemplate.Text;
                string signaturepath = tbSignature.Text;
                string savepath = tbWay.Text;

                await Task.Run(() =>
                {
                    this.Cursor = Cursors.WaitCursor;
                    btnCreate.Enabled = false;
                    buttonEdit.Enabled = false;
                    var inputData = new SystemInputData();
                     if (!inputData.ParseExcelFile(templatepath)) {
                        return;
                    };

                    var engine = new GeneratingReviews(inputData.generalData, inputData.students, inputData.advantages, inputData.disadvantages, signaturepath, savepath);
                    engine.CreateReviews();
                    MessageBox.Show(this, "Все готово!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);

                });



            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
                btnCreate.Enabled = true;
                buttonEdit.Enabled = true;
            }
            ;

        }

        private void buttonEdit_Click(object sender, EventArgs e)
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string originalTemplate = Path.Combine(baseDir, "revTemplate.docx");

            string tempTemplate = Path.Combine(Path.GetTempPath(), "template_temp.docx");

            try
            {

                if (!File.Exists(originalTemplate))
                {
                    MessageBox.Show($"Файл не найден по пути: {originalTemplate}");
                    return;
                }

                File.Copy(originalTemplate, tempTemplate, true);

                btnCreate.Enabled = false;
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = tempTemplate,
                    UseShellExecute = true
                };

                Process wordProcess = Process.Start(psi);

                MessageBox.Show("Шаблон открыт в Word. Отредактируйте его, сохраните, и закройте Word. Затем нажмите OK в этом окне", "Редактирование", MessageBoxButtons.OK);


                string[] requiredTags = {
                    "{NAME}", "{COURSE}", "{DIRECTIONOFTRAINING}",
                    "{DIRECTIVITY}", "{THEME}", "{ACADEMICTITLE}",
                    "{TEACHER}", "{GRADE}", "{DATE}",
                    "{ADVANTAGES}", "{DISADVANTAGES}",
                };
                bool tagNotContain = false;
                using (var doc = Xceed.Words.NET.DocX.Load(tempTemplate))
                {
                    string fullText = doc.Text;
                    foreach (var tag in requiredTags)
                    {
                        if (!fullText.Contains(tag))
                        {
                            MessageBox.Show($"Ошибка! Вы удалили обязательный тег {tag}. Изменения не будут сохранены", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            tagNotContain = true;
                        }
                    }
                }

                if (tagNotContain) return;

                File.Copy(tempTemplate, originalTemplate, true);

                MessageBox.Show("Шаблон успешно обновлен!", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Возникла ошибка при попытке редактирования {ex.Message}");
            }
            finally
            {
                if (File.Exists(tempTemplate)) File.Delete(tempTemplate);
                btnCreate.Enabled = true;
            }


        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            string helpFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "instruction.pdf");

            if (File.Exists(helpFilePath))
            {
                try
                {
                    Process.Start(new ProcessStartInfo(helpFilePath) { UseShellExecute = true });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Не удалось открыть файл инструкции: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
            }
            else {
                MessageBox.Show("Файл 'instruction.pdf' не найден в папке с программой.", "Файл отсутствует", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
