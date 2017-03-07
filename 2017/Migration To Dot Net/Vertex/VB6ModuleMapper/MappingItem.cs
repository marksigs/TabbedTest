using System;
using System.Collections.Generic;
using System.Text;

namespace VB6ModuleMapper
{
	class MappingItem
	{
		private string _vb6ModuleFileName;
		public string Vb6ModuleFileName
		{
			get { return _vb6ModuleFileName; }
			set { _vb6ModuleFileName = value; }
		}

		private string _directory;
		public string Directory
		{
			get { return _directory; }
			set { _directory = value; }
		}

		private string _csharpTypeName;
		public string CSharpTypeName
		{
			get { return _csharpTypeName; }
			set { _csharpTypeName = value; }
		}

		private string _obsoleteCSharpTypeName;
		public string ObsoleteCSharpTypeName
		{
			get { return _obsoleteCSharpTypeName; }
			set { _obsoleteCSharpTypeName = value; }
		}

		public MappingItem(string vb6ModuleFileName, string directory, string csharpTypeName)
			: this(vb6ModuleFileName, directory, csharpTypeName, null)
		{
		}

		public MappingItem(string vb6ModuleFileName, string directory, string csharpTypeName, string obsoleteCSharpTypeName)
		{
			_vb6ModuleFileName = vb6ModuleFileName;
			_directory = directory;
			_csharpTypeName = csharpTypeName;
			_obsoleteCSharpTypeName = obsoleteCSharpTypeName;
		}
	}
}
