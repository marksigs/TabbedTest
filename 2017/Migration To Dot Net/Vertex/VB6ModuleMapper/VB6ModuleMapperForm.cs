using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace VB6ModuleMapper
{
	public partial class VB6ModuleMapperForm : Form
	{
		public VB6ModuleMapperForm()
		{
			InitializeComponent();

			if (!string.IsNullOrEmpty(VB6ModuleMapper.Properties.Settings.Default.MappingFile))
			{
				textBoxMappingFile.Text = VB6ModuleMapper.Properties.Settings.Default.MappingFile;
			}
			if (!string.IsNullOrEmpty(VB6ModuleMapper.Properties.Settings.Default.RootDirectory))
			{
				textBoxRootDirectory.Text = VB6ModuleMapper.Properties.Settings.Default.RootDirectory;
			}

			UpdateControls();
		}

		private void buttonBrowseInputFile_Click(object sender, EventArgs e)
		{
			if (File.Exists(textBoxMappingFile.Text))
			{
				openFileDialogMappingFile.InitialDirectory = Path.GetDirectoryName(textBoxMappingFile.Text);
				openFileDialogMappingFile.FileName = Path.GetFileName(textBoxMappingFile.Text);
			}
			if (openFileDialogMappingFile.ShowDialog() == DialogResult.OK)
			{
				textBoxMappingFile.Text = openFileDialogMappingFile.FileName;
				UpdateControls();
			}
		}

		private void VB6ModuleMapperForm_FormClosed(object sender, FormClosedEventArgs e)
		{
			VB6ModuleMapper.Properties.Settings.Default.MappingFile = textBoxMappingFile.Text;
			VB6ModuleMapper.Properties.Settings.Default.RootDirectory = textBoxRootDirectory.Text;
			VB6ModuleMapper.Properties.Settings.Default.Save();
		}

		private void buttonExit_Click(object sender, EventArgs e)
		{
			Close();
		}

		private void buttonBrowseRootDirectory_Click(object sender, EventArgs e)
		{
			if (Directory.Exists(textBoxRootDirectory.Text))
			{
				folderBrowserDialogRootDirectory.SelectedPath = textBoxRootDirectory.Text;
			}
			if (folderBrowserDialogRootDirectory.ShowDialog() == DialogResult.OK)
			{
				textBoxRootDirectory.Text = folderBrowserDialogRootDirectory.SelectedPath;
				UpdateControls();
			}
		}

		private void buttonGenerate_Click(object sender, EventArgs e)
		{
			UpdateControls();
			if (!buttonGenerate.Enabled)
			{
				MessageBox.Show("Invalid mapping file or root directory", "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
			else
			{
				try
				{
					Analyse(textBoxMappingFile.Text, textBoxRootDirectory.Text);
				}
				catch (Exception exception)
				{
					MessageBox.Show(exception.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}
			}
		}

		private void UpdateControls()
		{
			buttonGenerate.Enabled = File.Exists(textBoxMappingFile.Text) && Directory.Exists(textBoxRootDirectory.Text);
		}

		private void textBoxMappingFile_TextChanged(object sender, EventArgs e)
		{
			UpdateControls();
		}

		private void textBoxRootDirectory_TextChanged(object sender, EventArgs e)
		{
			UpdateControls();
		}

		private void Analyse(string mappingFileName, string rootDirectory)
		{
			MappingFile inputMappingFile = new MappingFile(textBoxMappingFile.Text);
			inputMappingFile.Load();
			MappingFile outputMappingFile = new MappingFile(Path.ChangeExtension(textBoxMappingFile.Text, ".output.csv"));
			AnalyseDirectory(inputMappingFile, ref outputMappingFile, rootDirectory);
			outputMappingFile.Save();
		}

		private void AnalyseDirectory(MappingFile inputMappingFile, ref MappingFile outputMappingFile, string directory)
		{
			foreach (string file in Directory.GetFiles(directory))
			{
				string fileName = Path.GetFileName(file);
				string directoryName = Path.GetDirectoryName(file);
				if (inputMappingFile.Vb6ToCSharpMap.ContainsKey(fileName))
				{
					List<MappingItem> inputMappingItems = inputMappingFile.Vb6ToCSharpMap[fileName];
					if (inputMappingItems.Count > 0)
					{
						MappingItem inputMappingItem = inputMappingItems[0];
						List<MappingItem> outputMappingItems = null;
						if (outputMappingFile.Vb6ToCSharpMap.ContainsKey(fileName))
						{
							outputMappingItems = outputMappingFile.Vb6ToCSharpMap[fileName];
						}
						else
						{
							outputMappingItems = new List<MappingItem>();
							outputMappingFile.Vb6ToCSharpMap.Add(fileName, outputMappingItems);
						}
						outputMappingItems.Add(new MappingItem(fileName, directoryName, inputMappingItem.CSharpTypeName, inputMappingItem.ObsoleteCSharpTypeName));
					}
				}
			}

			foreach (string subDirectory in Directory.GetDirectories(directory))
			{
				AnalyseDirectory(inputMappingFile, ref outputMappingFile, subDirectory);
			}
		}
	}
}