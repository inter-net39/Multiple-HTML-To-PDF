using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace MultipleHtmlsToPDF
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private string[] PathTable = null;
        private string OutputPath = null;
        private string OutputFileName = null;

        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Multiselect = true;
                dialog.Filter = "HTML files (*.html)|*.html";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    PathTable = dialog.FileNames;
                    richTextBox1.Text = string.Join("\r\n", PathTable);
                    label7.Text = (dialog.FileNames.Length).ToString();

                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Process cmd = new Process();
            cmd.StartInfo.FileName = "cmd.exe";
            cmd.StartInfo.CreateNoWindow = false;
            cmd.StartInfo.RedirectStandardInput = true;
            cmd.StartInfo.RedirectStandardOutput = false;
            cmd.StartInfo.UseShellExecute = true;
            cmd.StartInfo.Verb = "runas";
            cmd.Start();

            cmd.StandardInput.WriteLine("cd C:\\Program Files\\wkhtmltopdf\\bin");
            cmd.StandardInput.Flush();

            int count = 0;
            foreach (var path in PathTable)
            {
                count++;
                cmd.StandardInput.WriteLine($"wkhtmltopdf \"{path}\" \"{OutputPath + "\\ConvertedPDF-" + count + ".pdf"}\"");
                cmd.StandardInput.Flush();
            }
            cmd.StandardInput.Close();
            MessageBox.Show("Done");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string path = System.AppDomain.CurrentDomain.BaseDirectory;
            path += "../../Instalation/wkhtmltox-0.12.5-1.msvc2015-win64.exe";
            if (File.Exists(path))
            {
                Process.Start(path);
            }
            else
            {
                MessageBox.Show("Instalation file not exist.");
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (File.Exists("C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe"))
            {
                label1.Text = "wkhtmltopdf is installed";
                label1.ForeColor = Color.DarkGreen;

            }
            else
            {
                label1.Text = "Install wkhtmltopdf first!";
                label1.ForeColor = Color.Red;
            }
            SetOutputPath(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));
            SetOutputFilenName("MyPDF");
        }

        private void SetOutputPath(string newPath)
        {
            OutputPath = newPath;
            label3.Text = newPath;
        }
        private void SetOutputFilenName(string newPath)
        {
            OutputFileName = newPath;
            label5.Text = newPath;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    SetOutputPath(fbd.SelectedPath);
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (textBox1.Visible)
            {
                if (!string.IsNullOrWhiteSpace(textBox1.Text))
                {
                    SetOutputFilenName(textBox1.Text);
                    textBox1.Text = "";
                }
                textBox1.Visible = false;
            }
            else
            {
                textBox1.Visible = true;
            }
        }
    }
}
