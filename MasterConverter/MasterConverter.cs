using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.WindowsAPICodePack.Dialogs;
using Microsoft.WindowsAPICodePack.Shell;

namespace MasterConverter
{
    public partial class MasterConverter : Form
    {
        public MasterConverter()
        {
            InitializeComponent();
            this.Text += "(Ver." + Assembly.GetEntryAssembly().GetName().Version + ")";
        }

        private void MasterConverter_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        private void MasterConverter_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files.Length == 1)
            {
                textBox1.Text = files[0];
            }
            else
            {
                MessageBox.Show("一つのファイルを入れて下さい。", Assembly.GetExecutingAssembly().GetName().Name, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            string filePath = OpenDlg("変換ファイル", "変換ファイル|*.*");
            if (string.IsNullOrEmpty(filePath)) return;
            textBox1.Text = filePath;
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            var csvWriter = new CsvWriter();
            csvWriter.ConvertFileToCsv(textBox1.Text);
        }

        private string OpenDlg(string title, string filter)
        {
            var dialog = new CommonOpenFileDialog
            {
                EnsurePathExists = true,
                EnsureFileExists = false,
                AllowNonFileSystemItems = false,
                DefaultFileName = "Select Folder",
                Title = "Select The Folder To Process"
            };
            return "";
        }
    }
}
