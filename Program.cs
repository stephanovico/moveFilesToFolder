using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace MoveToNewFolder
{
    internal static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
    
                if (args == null || args.Length == 0)
                {
                    MessageBox.Show("This application is designed to be used from SendTo.\n\nUsage:\n1. Select one or more files/folders\n2. Right-click → Send To → Move to New Folder",
                        "Move to New Folder", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                string folderName = PromptForFolderName();

                if (string.IsNullOrWhiteSpace(folderName))
                {
                    return;
                }

                string? parentDir = GetCommonParentDirectory(args);

                if (parentDir == null)
                {
                    MessageBox.Show("Selected files must share the same parent folder.",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                
                string targetDir = Path.Combine(parentDir, folderName);
                
                if (Directory.Exists(targetDir))
                {
                    var result = MessageBox.Show($"Folder '{folderName}' already exists. Move files into it?",
                        "Folder Exists", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result != DialogResult.Yes)
                    {
                        return;
                    }
                }
                else
                {
                    Directory.CreateDirectory(targetDir);
                }

                int moved = 0;
                int errors = 0;
                string errorstring = "";
                foreach (string path in args)
                {
                    try
                    {
                        string? name = Path.GetFileName(path);
                        if (string.IsNullOrEmpty(name))
                            continue;

                        string dest = Path.Combine(targetDir, name);

                        if (Directory.Exists(path))
                        {
                            if (dest != path) // Don't move into itself
                            {
                                Directory.Move(path, dest);
                                moved++;
                            }
                        }
                        else if (File.Exists(path))
                        {
                            File.Move(path, dest, overwrite: true);
                            moved++;
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error: {ex.Message}","Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                MessageBox.Show($"Moved {moved} files","Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}","Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static string PromptForFolderName()
        {
            string result = "";

            using var form = new Form
            {
                Width = 450,
                Height = 120,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "Move to New Folder",
                StartPosition = FormStartPosition.CenterScreen,
                MaximizeBox = false,
                MinimizeBox = false,
                TopMost = true,
                ShowInTaskbar = true
            };

            var textBox = new TextBox
            {
                Left = 15,
                Top = 20,
                Width = 400,
                Font = new System.Drawing.Font("Segoe UI", 9F)
            };

            // Handle Enter key
            textBox.KeyDown += (sender, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    result = textBox.Text;
                    form.DialogResult = DialogResult.OK;
                    form.Close();
                    e.SuppressKeyPress = true; // prevents beep
                }
            };

            form.Controls.Add(textBox);

            form.Shown += (s, e) => textBox.Focus();

            return form.ShowDialog() == DialogResult.OK ? result : "";
        }



        private static string? GetCommonParentDirectory(string[] paths)
        {
            var parentDirs = paths
                .Select(p => Path.GetDirectoryName(p))
                .Where(p => p != null)
                .Distinct()
                .ToList();

            return parentDirs.Count == 1 ? parentDirs[0] : null;
        }
    }
}