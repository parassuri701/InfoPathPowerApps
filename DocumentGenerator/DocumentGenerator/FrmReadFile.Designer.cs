namespace DocumentGenerator
{
    partial class FrmReadFile
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            bgWorkerRead = new System.ComponentModel.BackgroundWorker();
            btnXsn = new Button();
            colorDialog1 = new ColorDialog();
            pnlSingle = new Panel();
            label2 = new Label();
            label1 = new Label();
            btnOutput = new Button();
            txtOutputDirectory = new TextBox();
            lblInput = new Label();
            btnSelect = new Button();
            txtFileName = new TextBox();
            lblOutputDir = new Label();
            label3 = new Label();
            rbSingle = new RadioButton();
            rbMultiple = new RadioButton();
            pnlMultiple = new Panel();
            label4 = new Label();
            btnOutputDir = new Button();
            txtOutputDir = new TextBox();
            lblInputDir = new Label();
            btnSelectDir = new Button();
            txtInputDir = new TextBox();
            btnDirectory = new Button();
            richTextBox1 = new RichTextBox();
            button1 = new Button();
            pictureBox1 = new PictureBox();
            pnlSingle.SuspendLayout();
            pnlMultiple.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)pictureBox1).BeginInit();
            SuspendLayout();
            // 
            // bgWorkerRead
            // 
            bgWorkerRead.WorkerReportsProgress = true;
            // 
            // btnXsn
            // 
            btnXsn.Location = new Point(425, 786);
            btnXsn.Margin = new Padding(4, 5, 4, 5);
            btnXsn.Name = "btnXsn";
            btnXsn.Size = new Size(271, 55);
            btnXsn.TabIndex = 7;
            btnXsn.Text = "Clear";
            btnXsn.UseVisualStyleBackColor = true;
            btnXsn.Click += btnXsn_Click;
            // 
            // pnlSingle
            // 
            pnlSingle.Controls.Add(label2);
            pnlSingle.Controls.Add(label1);
            pnlSingle.Controls.Add(btnOutput);
            pnlSingle.Controls.Add(txtOutputDirectory);
            pnlSingle.Controls.Add(lblInput);
            pnlSingle.Controls.Add(btnSelect);
            pnlSingle.Controls.Add(txtFileName);
            pnlSingle.Location = new Point(1, 282);
            pnlSingle.Margin = new Padding(4, 5, 4, 5);
            pnlSingle.Name = "pnlSingle";
            pnlSingle.Size = new Size(1112, 138);
            pnlSingle.TabIndex = 12;
            pnlSingle.Visible = false;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(476, 27);
            label2.Margin = new Padding(4, 0, 4, 0);
            label2.Name = "label2";
            label2.Size = new Size(0, 25);
            label2.TabIndex = 18;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(17, 86);
            label1.Margin = new Padding(4, 0, 4, 0);
            label1.Name = "label1";
            label1.Size = new Size(146, 25);
            label1.TabIndex = 17;
            label1.Text = "Output Directory";
            // 
            // btnOutput
            // 
            btnOutput.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnOutput.Location = new Point(1037, 77);
            btnOutput.Margin = new Padding(4, 5, 4, 5);
            btnOutput.Name = "btnOutput";
            btnOutput.Size = new Size(66, 40);
            btnOutput.TabIndex = 16;
            btnOutput.Text = ". . .";
            btnOutput.UseVisualStyleBackColor = true;
            btnOutput.Click += btnOutput_Click;
            // 
            // txtOutputDirectory
            // 
            txtOutputDirectory.BackColor = SystemColors.ControlLightLight;
            txtOutputDirectory.Location = new Point(163, 81);
            txtOutputDirectory.Margin = new Padding(4, 5, 4, 5);
            txtOutputDirectory.Name = "txtOutputDirectory";
            txtOutputDirectory.ReadOnly = true;
            txtOutputDirectory.Size = new Size(871, 31);
            txtOutputDirectory.TabIndex = 15;
            // 
            // lblInput
            // 
            lblInput.AutoSize = true;
            lblInput.Location = new Point(17, 27);
            lblInput.Margin = new Padding(4, 0, 4, 0);
            lblInput.Name = "lblInput";
            lblInput.Size = new Size(116, 25);
            lblInput.TabIndex = 14;
            lblInput.Text = "XSN/NFP File";
            // 
            // btnSelect
            // 
            btnSelect.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnSelect.Location = new Point(1037, 18);
            btnSelect.Margin = new Padding(4, 5, 4, 5);
            btnSelect.Name = "btnSelect";
            btnSelect.Size = new Size(66, 40);
            btnSelect.TabIndex = 13;
            btnSelect.Text = ". . .";
            btnSelect.UseVisualStyleBackColor = true;
            btnSelect.Click += btnSelect_Click;
            // 
            // txtFileName
            // 
            txtFileName.BackColor = SystemColors.ControlLightLight;
            txtFileName.Location = new Point(163, 22);
            txtFileName.Margin = new Padding(4, 5, 4, 5);
            txtFileName.Name = "txtFileName";
            txtFileName.ReadOnly = true;
            txtFileName.Size = new Size(871, 31);
            txtFileName.TabIndex = 12;
            // 
            // lblOutputDir
            // 
            lblOutputDir.AutoSize = true;
            lblOutputDir.Location = new Point(12, 77);
            lblOutputDir.Margin = new Padding(4, 0, 4, 0);
            lblOutputDir.Name = "lblOutputDir";
            lblOutputDir.Size = new Size(146, 25);
            lblOutputDir.TabIndex = 17;
            lblOutputDir.Text = "Output Directory";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Font = new Font("Segoe UI", 24F, FontStyle.Bold | FontStyle.Underline);
            label3.ForeColor = SystemColors.Highlight;
            label3.Location = new Point(105, 158);
            label3.Margin = new Padding(4, 0, 4, 0);
            label3.Name = "label3";
            label3.Size = new Size(985, 65);
            label3.TabIndex = 13;
            label3.Text = "InfoPath and Nintex Document Generator ";
            // 
            // rbSingle
            // 
            rbSingle.AutoSize = true;
            rbSingle.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            rbSingle.Location = new Point(306, 239);
            rbSingle.Margin = new Padding(4, 5, 4, 5);
            rbSingle.Name = "rbSingle";
            rbSingle.Size = new Size(155, 36);
            rbSingle.TabIndex = 16;
            rbSingle.TabStop = true;
            rbSingle.Text = "Single File";
            rbSingle.UseVisualStyleBackColor = true;
            rbSingle.CheckedChanged += rbSingle_CheckedChanged;
            // 
            // rbMultiple
            // 
            rbMultiple.AutoSize = true;
            rbMultiple.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            rbMultiple.Location = new Point(482, 239);
            rbMultiple.Margin = new Padding(4, 5, 4, 5);
            rbMultiple.Name = "rbMultiple";
            rbMultiple.Size = new Size(192, 36);
            rbMultiple.TabIndex = 17;
            rbMultiple.TabStop = true;
            rbMultiple.Text = "Multiple Files";
            rbMultiple.UseVisualStyleBackColor = true;
            rbMultiple.CheckedChanged += rbMultiple_CheckedChanged;
            // 
            // pnlMultiple
            // 
            pnlMultiple.Controls.Add(label4);
            pnlMultiple.Controls.Add(lblOutputDir);
            pnlMultiple.Controls.Add(btnOutputDir);
            pnlMultiple.Controls.Add(txtOutputDir);
            pnlMultiple.Controls.Add(lblInputDir);
            pnlMultiple.Controls.Add(btnSelectDir);
            pnlMultiple.Controls.Add(txtInputDir);
            pnlMultiple.Location = new Point(1, 285);
            pnlMultiple.Margin = new Padding(4, 5, 4, 5);
            pnlMultiple.Name = "pnlMultiple";
            pnlMultiple.Size = new Size(1112, 140);
            pnlMultiple.TabIndex = 19;
            pnlMultiple.Visible = false;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(476, 27);
            label4.Margin = new Padding(4, 0, 4, 0);
            label4.Name = "label4";
            label4.Size = new Size(0, 25);
            label4.TabIndex = 18;
            // 
            // btnOutputDir
            // 
            btnOutputDir.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnOutputDir.Location = new Point(1042, 72);
            btnOutputDir.Margin = new Padding(4, 5, 4, 5);
            btnOutputDir.Name = "btnOutputDir";
            btnOutputDir.Size = new Size(66, 40);
            btnOutputDir.TabIndex = 16;
            btnOutputDir.Text = ". . .";
            btnOutputDir.UseVisualStyleBackColor = true;
            btnOutputDir.Click += btnOutputDir_Click;
            // 
            // txtOutputDir
            // 
            txtOutputDir.BackColor = SystemColors.ControlLightLight;
            txtOutputDir.Location = new Point(166, 74);
            txtOutputDir.Margin = new Padding(4, 5, 4, 5);
            txtOutputDir.Name = "txtOutputDir";
            txtOutputDir.ReadOnly = true;
            txtOutputDir.Size = new Size(871, 31);
            txtOutputDir.TabIndex = 15;
            // 
            // lblInputDir
            // 
            lblInputDir.AutoSize = true;
            lblInputDir.Location = new Point(12, 18);
            lblInputDir.Margin = new Padding(4, 0, 4, 0);
            lblInputDir.Name = "lblInputDir";
            lblInputDir.Size = new Size(131, 25);
            lblInputDir.TabIndex = 14;
            lblInputDir.Text = "Input Directory";
            // 
            // btnSelectDir
            // 
            btnSelectDir.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnSelectDir.Location = new Point(1042, 14);
            btnSelectDir.Margin = new Padding(4, 5, 4, 5);
            btnSelectDir.Name = "btnSelectDir";
            btnSelectDir.Size = new Size(66, 40);
            btnSelectDir.TabIndex = 13;
            btnSelectDir.Text = ". . .";
            btnSelectDir.UseVisualStyleBackColor = true;
            btnSelectDir.Click += btnSelectDir_Click;
            // 
            // txtInputDir
            // 
            txtInputDir.BackColor = SystemColors.ControlLightLight;
            txtInputDir.Location = new Point(166, 15);
            txtInputDir.Margin = new Padding(4, 5, 4, 5);
            txtInputDir.Name = "txtInputDir";
            txtInputDir.ReadOnly = true;
            txtInputDir.Size = new Size(871, 31);
            txtInputDir.TabIndex = 12;
            txtInputDir.TextChanged += txtInputDir_TextChanged;
            // 
            // btnDirectory
            // 
            btnDirectory.Location = new Point(58, 786);
            btnDirectory.Margin = new Padding(4, 5, 4, 5);
            btnDirectory.Name = "btnDirectory";
            btnDirectory.Size = new Size(271, 55);
            btnDirectory.TabIndex = 19;
            btnDirectory.Text = "Generate Documentation";
            btnDirectory.UseVisualStyleBackColor = true;
            btnDirectory.Click += btnDirectory_Click;
            // 
            // richTextBox1
            // 
            richTextBox1.Location = new Point(0, 421);
            richTextBox1.Name = "richTextBox1";
            richTextBox1.Size = new Size(1109, 355);
            richTextBox1.TabIndex = 21;
            richTextBox1.Text = "";
            // 
            // button1
            // 
            button1.Location = new Point(806, 786);
            button1.Name = "button1";
            button1.Size = new Size(284, 55);
            button1.TabIndex = 22;
            button1.Text = "Cancel";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // pictureBox1
            // 
            pictureBox1.Image = Properties.Resources.yash_background;
            pictureBox1.Location = new Point(0, 0);
            pictureBox1.Name = "pictureBox1";
            pictureBox1.Size = new Size(1122, 155);
            pictureBox1.TabIndex = 23;
            pictureBox1.TabStop = false;
            // 
            // FrmReadFile
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1122, 855);
            Controls.Add(pictureBox1);
            Controls.Add(button1);
            Controls.Add(pnlSingle);
            Controls.Add(btnDirectory);
            Controls.Add(richTextBox1);
            Controls.Add(pnlMultiple);
            Controls.Add(rbSingle);
            Controls.Add(btnXsn);
            Controls.Add(rbMultiple);
            Controls.Add(label3);
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            Margin = new Padding(4, 5, 4, 5);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "FrmReadFile";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Document Generator";
            Load += FrmReadFile_Load;
            pnlSingle.ResumeLayout(false);
            pnlSingle.PerformLayout();
            pnlMultiple.ResumeLayout(false);
            pnlMultiple.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)pictureBox1).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private System.ComponentModel.BackgroundWorker bgWorkerRead;
        private Button btnXsn;
        private ColorDialog colorDialog1;
        private Panel pnlSingle;
        private Label label2;
        private Label label1;
        private Button btnOutput;
        private Label lblInput;
        private Button btnSelect;
        private TextBox txtFileName;
        private Label label3;
        private RadioButton rbSingle;
        private RadioButton rbMultiple;
        private Panel pnlMultiple;
        private Label label4;
        private Label lblOutputDir;
        private Button btnOutputDir;
        private TextBox txtOutputDir;
        private Label lblInputDir;
        private Button btnSelectDir;
        private TextBox txtInputDir;
        private TextBox txtOutputDirectory;
        private Button btnDirectory;
        private RichTextBox richTextBox1;
        private Button button1;
        private PictureBox pictureBox1;
    }
}