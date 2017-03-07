using System;
using System.Collections.Generic;
using System.Text;

namespace PreMigrationConverter
{
    internal class ConverterEventArgs : EventArgs
    {
        private string _fileName;
        private bool _fileAmended;

        public bool FileAmended
        {
            get { return _fileAmended; }
            set { _fileAmended = value; }
        }

        public string FileName
        {
            get { return _fileName; }
            set { _fileName = value; }
        }
    }
}
