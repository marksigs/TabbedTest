using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace PreMigrationConverter
{
    /// <summary>
    /// Handle conversion 
    /// </summary>
    internal class Converter
    {
        private string _inputPath; 
        private TextWriter _log;
        private bool _cancelled;
        private PreMigrationConverter.frmMain.ConverterFinishedCallback _callback;
        private List<FileSystemInfo> _fileInfoList;

        public Converter()
        {
            _fileInfoList = new List<FileSystemInfo>();
            _cancelled = false;
        }

        /// <summary>
        /// Initialise the converter
        /// </summary>
        /// <param name="inputPath"></param>
        /// <param name="logFileName"></param>
        /// <param name="callback"></param>
        public void Initialise(
            string inputPath, 
            string logFileName,
            PreMigrationConverter.frmMain.ConverterFinishedCallback callback)
        {
            if (inputPath == null || logFileName == null)
            {
                throw new ArgumentNullException();
            }
            _inputPath = inputPath;
            _log = File.CreateText(logFileName);
            _callback = callback;
        }

        /// <summary>
        /// Generate a private list of files to be converted
        /// </summary>
        /// <param name="path"></param>
        private void BuildFileList(string path)
        {
            DirectoryInfo di = new DirectoryInfo(path);

            foreach (FileInfo fi in di.GetFiles())
            {
                string extension = fi.Extension;
                if (String.Compare(extension, ".cls") == 0 || 
                    String.Compare(extension, ".bas") == 0)
                {
                    _fileInfoList.Add(fi);
                }
            }
            foreach (DirectoryInfo subdirectory in di.GetDirectories())
            {
                BuildFileList(subdirectory.FullName);
            }
        }

        /// <summary>
        /// Handle finished converting a file event 
        /// </summary>
        /// <remarks>Handle event by reraising it</remarks>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void  _folderConverter_FinishedConvertingFile(object sender, ConverterEventArgs e)
        {
            OnFinishedConvertingFile(e);
        }

        /// <summary>
        /// Do the conversion
        /// </summary>
        public void Start()
        {
            _fileInfoList.Clear();  //GHun
            BuildFileList(_inputPath);
            bool success = true;

            try
            {
                foreach (FileInfo fi in _fileInfoList)
                {
                    if (success && !_cancelled)
                    {
                        string inFileName = fi.FullName;
                        string outFileName = inFileName + "_out";

                        FileConverter fc = new FileConverter(
                            inFileName, outFileName, _log);
                        success = fc.Convert();
                        ConverterEventArgs e = new ConverterEventArgs();
                        e.FileAmended = fc.FileAmended;
                        e.FileName = inFileName;
                        OnFinishedConvertingFile(e);
                    }
                }
            }
            catch(Exception ex)
            {
                _log.WriteLine("Error: " + ex.Message.ToString());
                _log.WriteLine(ex.StackTrace);
                success = false;
            }
            finally
            {
                _log.Flush();
                _log.Close();
                _callback.Invoke(success);
            }
        }

        public delegate void ConverterEventHandler(object sender, ConverterEventArgs e);

        public event ConverterEventHandler FinishedConvertingFile;

        /// <summary>
        /// Raise event
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnFinishedConvertingFile(ConverterEventArgs e)
        {
            if (FinishedConvertingFile != null)
            {
                FinishedConvertingFile(this, e);
            }
        }
    }
}
