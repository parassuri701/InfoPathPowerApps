using DocumentGeneratorService;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DocumentGenerator
{
    public partial class FrmReadFile : Form
    {
        public FrmReadFile()
        {
            InitializeComponent();
            Bitmap originalImage = new Bitmap(pictureBox1.Image);
            Bitmap filledImage = new Bitmap(1122, 155);

            using (Graphics g = Graphics.FromImage(filledImage))
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;

                // Determine the aspect ratio of the original image
                float originalAspect = (float)originalImage.Width / originalImage.Height;
                float targetAspect = (float)1122 / 155;

                int cropWidth, cropHeight;
                if (originalAspect > targetAspect)
                {
                    // Crop width to match target aspect ratio
                    cropWidth = (int)(originalImage.Height * targetAspect);
                    cropHeight = originalImage.Height;
                }
                else
                {
                    // Crop height to match target aspect ratio
                    cropWidth = originalImage.Width;
                    cropHeight = (int)(originalImage.Width / targetAspect);
                }

                int cropX = (originalImage.Width - cropWidth) / 2;
                int cropY = (originalImage.Height - cropHeight) / 2;

                g.DrawImage(originalImage, new Rectangle(0, 0, 1122, 155), new Rectangle(cropX, cropY, cropWidth, cropHeight), GraphicsUnit.Pixel);
            }

            pictureBox1.Image = filledImage;
        }

        private void btnXsn_Click(object sender, EventArgs e)
        {
            this.txtFileName.Text = "";
            this.txtInputDir.Text = "";
            this.txtOutputDir.Text = "";
            this.txtOutputDirectory.Text = "";
            this.richTextBox1.Text = "";
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "*.xsn;*.nfp files|*.xsn;*.nfp"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtFileName.Text = openFileDialog.FileName;
            }
        }

        private void btnOutput_Click(object sender, EventArgs e)
        {
            // Create a new FolderBrowserDialog instance
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                // Optionally set properties, like the starting directory
                folderDialog.SelectedPath = @"C:\"; // Or any other default path

                // Show the FolderBrowserDialog and check if the user selected a folder
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    string directory = folderDialog.SelectedPath;
                    string fileName = Path.GetFileNameWithoutExtension(txtFileName.Text);
                    string docFilePath = Path.Combine(directory, fileName + ".docx");
                    // Get the selected folder path
                    txtOutputDirectory.Text = docFilePath;
                }
            }
        }

        private void rbSingle_CheckedChanged(object sender, EventArgs e)
        {
            pnlSingle.Visible = rbSingle.Checked;
            pnlMultiple.Visible = !rbSingle.Checked;
            this.richTextBox1.Text = "";
        }

        private void rbMultiple_CheckedChanged(object sender, EventArgs e)
        {
            pnlSingle.Visible = !rbMultiple.Checked;
            pnlMultiple.Visible = rbMultiple.Checked;
            this.richTextBox1.Text = "";
        }

        private void btnSelectDir_Click(object sender, EventArgs e)
        {
            // Create a new FolderBrowserDialog instance
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Reset();
                // Optionally set properties, like the starting directory
                folderDialog.SelectedPath = @"C:\"; // Or any other default path

                // Show the FolderBrowserDialog and check if the user selected a folder
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    // Get the selected folder path
                    txtInputDir.Text = folderDialog.SelectedPath;
                }
            }
        }

        private void btnOutputDir_Click(object sender, EventArgs e)
        {
            // Create a new FolderBrowserDialog instance
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Reset();
                // Optionally set properties, like the starting directory
                folderDialog.SelectedPath = @"C:\"; // Or any other default path

                // Show the FolderBrowserDialog and check if the user selected a folder
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    // Get the selected folder path
                    txtOutputDir.Text = folderDialog.SelectedPath;
                }
            }
        }

        private void btnDirectory_Click(object sender, EventArgs e)
        {
            if (rbSingle.Checked)
            {
                try
                {
                    string xsnFilePath = txtFileName.Text;
                    string docFilePath = txtOutputDirectory.Text;
                    string outputFolderPath = Path.GetDirectoryName(docFilePath);
                    // Extract InfoPath form contents and generate documentation
                    //InfoPathGenerateDocument generator = new InfoPathGenerateDocument();
                    IGenerateDocument generator = DocumentGeneratoryFactory.GetDocumentGeneratorByType(xsnFilePath);
                    generator.GenerateDocumentation(xsnFilePath, docFilePath);
                    this.richTextBox1.Text = RemoveUnwantedString(generator.Output);
                    MessageBox.Show($"Documentation successfully generated at {docFilePath}");

                    Process.Start("explorer.exe", Path.GetDirectoryName(docFilePath));
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"An error occurred: {ex.Message}");
                    Console.WriteLine(ex.StackTrace);
                }
            }
            else if (rbMultiple.Checked)
            {
                string output = string.Empty;
                string directoryPath = txtInputDir.Text;

                // Check if the directory exists
                if (Directory.Exists(directoryPath))
                {
                    // Get all .xsn files from the directory (including subdirectories)
                    string[] xsnFiles = Directory.GetFiles(directoryPath, "*.xsn", SearchOption.AllDirectories)
                             .Concat(Directory.GetFiles(directoryPath, "*.nfp", SearchOption.AllDirectories))
                             .ToArray();

                    // If there are any files, print the filenames
                    if (xsnFiles.Length > 0)
                    {
                        foreach (string file in xsnFiles)
                        {
                            string xsnFilePath = Path.GetFileName(file);
                            string fileName = Path.GetFileNameWithoutExtension(xsnFilePath);
                            string docFilePath = Path.Combine(txtOutputDir.Text, fileName + ".docx");
                            // Extract InfoPath form contents and generate documentation
                            //InfoPathGenerateDocument generator = new InfoPathGenerateDocument();
                            IGenerateDocument generator = DocumentGeneratoryFactory.GetDocumentGeneratorByType(file);
                            generator.GenerateDocumentation(file, docFilePath);

                            output += generator.Output +"\n\n-----------------------------------\n\n";
                        }
                        this.richTextBox1.Text = RemoveUnwantedString(output);
                        MessageBox.Show($"Documentation successfully generated at {txtOutputDir.Text}");
                        Process.Start("explorer.exe", txtOutputDir.Text);
                    }
                    else
                    {
                        MessageBox.Show("No .xsn files found in the specified directory.");
                    }
                }
            }

        }

        private string RemoveUnwantedString(string output)
        {
            // Specify the strings to remove
            string[] unwantedStrings = {
            "Microsoft (R) File Expansion Utility",
            "Copyright (c) Microsoft Corporation. All rights reserved.",
            "to Extraction Queue",
            "Expanding Files ....",
            "Expanding Files Complete ..."
        };

            // Process the input string
            foreach (var unwanted in unwantedStrings)
            {
                output = output.Replace(unwanted, string.Empty);
            }

            var lines = output.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            var cleanedLines = lines
                .Select((line, index) =>
                {
                    // Keep non-blank lines or single blank lines (not consecutive)
                    if (!string.IsNullOrWhiteSpace(line))
                    {
                        return line.Trim(); // Non-blank line
                    }
                    else if (index == 0 || !string.IsNullOrWhiteSpace(lines[index - 1]))
                    {
                        return string.Empty; // Preserve single blank line
                    }
                    return null; // Remove consecutive blank lines
                })
                .Where(line => line != null) // Exclude null lines
                .ToList();

            // Join the lines back with single-line spacing
            output = string.Join(Environment.NewLine, cleanedLines);
            return output;
        }

        private void txtInputDir_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FrmReadFile_Load(object sender, EventArgs e)
        {
            this.rbSingle.Checked = true;
        }
    }
}
