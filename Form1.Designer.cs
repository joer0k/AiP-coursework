namespace coursework
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            label1 = new Label();
            pictureBox1 = new PictureBox();
            btnTemplate = new Button();
            btnSignature = new Button();
            btnWay = new Button();
            tbTemplate = new TextBox();
            tbSignature = new TextBox();
            tbWay = new TextBox();
            btnCreate = new Button();
            openTemplateDialog = new OpenFileDialog();
            openSignatureDialog = new OpenFileDialog();
            folderBrowserDialog = new FolderBrowserDialog();
            buttonEdit = new Button();
            toolStrip1 = new ToolStrip();
            toolStripButton1 = new ToolStripButton();
            ((System.ComponentModel.ISupportInitialize)pictureBox1).BeginInit();
            toolStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Cambria", 18F, FontStyle.Bold, GraphicsUnit.Point, 204);
            label1.Location = new Point(50, 43);
            label1.Name = "label1";
            label1.Size = new Size(239, 28);
            label1.TabIndex = 0;
            label1.Text = "Создание рецензий";
            // 
            // pictureBox1
            // 
            pictureBox1.Image = (Image)resources.GetObject("pictureBox1.Image");
            pictureBox1.Location = new Point(455, 28);
            pictureBox1.Name = "pictureBox1";
            pictureBox1.Size = new Size(180, 74);
            pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox1.TabIndex = 1;
            pictureBox1.TabStop = false;
            // 
            // btnTemplate
            // 
            btnTemplate.BackColor = Color.WhiteSmoke;
            btnTemplate.Font = new Font("Cambria", 9.75F, FontStyle.Bold);
            btnTemplate.Location = new Point(88, 119);
            btnTemplate.Name = "btnTemplate";
            btnTemplate.Size = new Size(146, 55);
            btnTemplate.TabIndex = 2;
            btnTemplate.Text = "Выбрать шаблон";
            btnTemplate.UseVisualStyleBackColor = false;
            btnTemplate.Click += btnTemplate_Click;
            // 
            // btnSignature
            // 
            btnSignature.BackColor = Color.WhiteSmoke;
            btnSignature.Font = new Font("Cambria", 9.75F, FontStyle.Bold);
            btnSignature.Location = new Point(88, 210);
            btnSignature.Name = "btnSignature";
            btnSignature.Size = new Size(146, 55);
            btnSignature.TabIndex = 3;
            btnSignature.Text = "Выбрать подпись";
            btnSignature.UseVisualStyleBackColor = false;
            btnSignature.Click += btnSignature_Click;
            // 
            // btnWay
            // 
            btnWay.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnWay.BackColor = Color.WhiteSmoke;
            btnWay.Font = new Font("Cambria", 9.75F, FontStyle.Bold);
            btnWay.Location = new Point(88, 308);
            btnWay.Name = "btnWay";
            btnWay.Size = new Size(146, 55);
            btnWay.TabIndex = 4;
            btnWay.Text = "Выбрать путь";
            btnWay.UseVisualStyleBackColor = false;
            btnWay.Click += btnWay_Click;
            // 
            // tbTemplate
            // 
            tbTemplate.BackColor = SystemColors.InactiveCaption;
            tbTemplate.Location = new Point(257, 136);
            tbTemplate.Name = "tbTemplate";
            tbTemplate.PlaceholderText = "Путь к необходимому шаблону";
            tbTemplate.ReadOnly = true;
            tbTemplate.Size = new Size(366, 23);
            tbTemplate.TabIndex = 5;
            // 
            // tbSignature
            // 
            tbSignature.BackColor = SystemColors.InactiveCaption;
            tbSignature.Location = new Point(257, 227);
            tbSignature.Name = "tbSignature";
            tbSignature.PlaceholderText = "Путь к необходимой подписи преподавателя";
            tbSignature.ReadOnly = true;
            tbSignature.Size = new Size(366, 23);
            tbSignature.TabIndex = 6;
            // 
            // tbWay
            // 
            tbWay.BackColor = SystemColors.InactiveCaption;
            tbWay.Location = new Point(257, 325);
            tbWay.Name = "tbWay";
            tbWay.PlaceholderText = "Путь для сохранения готовой рецензии";
            tbWay.ReadOnly = true;
            tbWay.Size = new Size(366, 23);
            tbWay.TabIndex = 7;
            // 
            // btnCreate
            // 
            btnCreate.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnCreate.BackColor = Color.WhiteSmoke;
            btnCreate.Font = new Font("Cambria", 9.75F, FontStyle.Bold);
            btnCreate.Location = new Point(88, 394);
            btnCreate.Name = "btnCreate";
            btnCreate.Size = new Size(146, 38);
            btnCreate.TabIndex = 8;
            btnCreate.Text = "Создать";
            btnCreate.UseVisualStyleBackColor = false;
            btnCreate.Click += btnCreate_Click;
            // 
            // openSignatureDialog
            // 
            openSignatureDialog.FileName = "openFileDialog2";
            // 
            // buttonEdit
            // 
            buttonEdit.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            buttonEdit.BackColor = Color.WhiteSmoke;
            buttonEdit.Font = new Font("Cambria", 9.75F, FontStyle.Bold);
            buttonEdit.Location = new Point(476, 394);
            buttonEdit.Name = "buttonEdit";
            buttonEdit.Size = new Size(147, 38);
            buttonEdit.TabIndex = 9;
            buttonEdit.Text = "Изменить исходный шаблон рецензии";
            buttonEdit.UseVisualStyleBackColor = false;
            buttonEdit.Click += buttonEdit_Click;
            // 
            // toolStrip1
            // 
            toolStrip1.BackColor = Color.White;
            toolStrip1.Items.AddRange(new ToolStripItem[] { toolStripButton1 });
            toolStrip1.Location = new Point(0, 0);
            toolStrip1.Name = "toolStrip1";
            toolStrip1.Size = new Size(685, 25);
            toolStrip1.TabIndex = 10;
            toolStrip1.Text = "toolStrip1";
            // 
            // toolStripButton1
            // 
            toolStripButton1.DisplayStyle = ToolStripItemDisplayStyle.Text;
            toolStripButton1.Image = (Image)resources.GetObject("toolStripButton1.Image");
            toolStripButton1.ImageTransparentColor = Color.Magenta;
            toolStripButton1.Name = "toolStripButton1";
            toolStripButton1.RightToLeft = RightToLeft.Yes;
            toolStripButton1.Size = new Size(23, 22);
            toolStripButton1.Text = "?";
            toolStripButton1.Click += toolStripButton1_Click;
            // 
            // Form1
            // 
            AutoScaleMode = AutoScaleMode.None;
            BackColor = Color.White;
            ClientSize = new Size(685, 495);
            Controls.Add(toolStrip1);
            Controls.Add(buttonEdit);
            Controls.Add(btnCreate);
            Controls.Add(tbWay);
            Controls.Add(tbSignature);
            Controls.Add(tbTemplate);
            Controls.Add(btnWay);
            Controls.Add(btnSignature);
            Controls.Add(btnTemplate);
            Controls.Add(pictureBox1);
            Controls.Add(label1);
            HelpButton = true;
            Icon = (Icon)resources.GetObject("$this.Icon");
            MaximizeBox = false;
            MaximumSize = new Size(701, 534);
            MinimizeBox = false;
            MinimumSize = new Size(701, 534);
            Name = "Form1";
            Text = "Генератор рецензий";
            ((System.ComponentModel.ISupportInitialize)pictureBox1).EndInit();
            toolStrip1.ResumeLayout(false);
            toolStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label label1;
        private PictureBox pictureBox1;
        private Button btnTemplate;
        private Button btnSignature;
        private Button btnWay;
        private TextBox tbTemplate;
        private TextBox tbSignature;
        private TextBox tbWay;
        private Button btnCreate;
        private OpenFileDialog openTemplateDialog;
        private OpenFileDialog openSignatureDialog;
        private FolderBrowserDialog folderBrowserDialog;
        private Button buttonEdit;
        private ToolStrip toolStrip1;
        private ToolStripButton toolStripButton1;
    }
}
