using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace VB6ModuleMapper
{
	class MappingFile
	{
		private string _fileName;
		public string FileName
		{
			get { return _fileName; }
		}

		private Dictionary<string, List<MappingItem>> _vb6ToCSharpMap = new Dictionary<string, List<MappingItem>>();
		internal Dictionary<string, List<MappingItem>> Vb6ToCSharpMap
		{
			get { return _vb6ToCSharpMap; }
			set { _vb6ToCSharpMap = value; }
		}

		public MappingFile(string fileName)
		{
			_fileName = fileName;
		}

		public void Load()
		{
			_vb6ToCSharpMap.Clear();

			int lineIndex = 0;
			using (StreamReader streamReader = new StreamReader(_fileName))
			{
				string line;
				while ((line = streamReader.ReadLine()) != null)
				{
					lineIndex++;
					string[] items = line.Split(',');

					if (items.Length == 4)
					{
						string fileName = items[0].Trim();
						List<MappingItem> mappingItems = null;
						if (_vb6ToCSharpMap.ContainsKey(fileName))
						{
							mappingItems = _vb6ToCSharpMap[fileName];
						}
						else
						{
							mappingItems = new List<MappingItem>();
							_vb6ToCSharpMap.Add(fileName, mappingItems);
						}
						mappingItems.Add(new MappingItem(fileName, items[1].Trim(), items[2].Trim(), items[3].Trim()));
					}
					else
					{
						throw new InvalidOperationException("Invalid line " + lineIndex.ToString() + ": " + line);
					}
				}
			}
		}

		public void Save()
		{
			using (StreamWriter streamWriter = new StreamWriter(_fileName))
			{
				streamWriter.WriteLine("VB6 Module,Directory,C# Type,Obsolete C# Type");
				foreach (KeyValuePair<string, List<MappingItem>> item in _vb6ToCSharpMap)
				{
					List<MappingItem> mappingItems = item.Value;
					foreach (MappingItem mappingItem in mappingItems)
					{
						streamWriter.WriteLine(mappingItem.Vb6ModuleFileName + "," + mappingItem.Directory + "," + mappingItem.CSharpTypeName + "," + mappingItem.ObsoleteCSharpTypeName);
					}
				}
			}
		}
	}
}
