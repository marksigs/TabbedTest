using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Xml;

namespace Vertex.Fsd.Omiga.VisualBasicPort
{
	public sealed class ResumeNext : IDisposable
	{
		private List<int> _omigaErrors = new List<int>();
		private XmlElement _xmlElement;

		private bool _disposed;

		public ResumeNext(params object[] parameters)
		{
			foreach (object parameter in parameters)
			{
				if (parameter is OMIGAERROR || parameter is Int32)
				{
					_omigaErrors.Add((int)parameter);
				}
				else if (parameter is XmlElement && _xmlElement == null)
				{
					_xmlElement = (XmlElement)parameter;
				}
			}

			if (_omigaErrors.Count > 0)
			{
				AddSuppressedErrors(_omigaErrors);
			}

			if (_xmlElement != null)
			{
				PushResumeOnWarningElement(_xmlElement);
			}
		}

		~ResumeNext()
		{
			Dispose(false);
		}

		private const string _suppressedErrorsDataSlotName = "SuppressedErrors";
		private const string _resumeOnWarningElementStackDataSlotName = "ResumeOnWarningElement";

		public static bool IsResumeNext(ErrAssistException exception)
		{
			return IsResumeNextError(exception) || IsResumeNextWarning(exception);
		}

		private static bool IsResumeNextError(ErrAssistException exception)
		{
			bool isResumeNextError = false;
			Dictionary<int, int> suppressedErrors = (Dictionary<int, int>)Thread.GetData(Thread.GetNamedDataSlot(_suppressedErrorsDataSlotName));
			if (suppressedErrors != null)
			{
				isResumeNextError = suppressedErrors.ContainsKey(exception.OmigaErrorNumber);
			}
			return isResumeNextError;
		}

		private static bool IsResumeNextWarning(ErrAssistException exception)
		{
			bool isResumeNextWarning = false;
			XmlElement xmlElement = null;
			if (exception.IsWarning() && (xmlElement = PeekResumeOnWarningElement()) != null)
			{
				exception.AddWarning(xmlElement);
				isResumeNextWarning = true;
			}
			return isResumeNextWarning;
		}

		private static void AddSuppressedErrors(List<int> omigaErrors)
		{
			LocalDataStoreSlot slot = Thread.GetNamedDataSlot(_suppressedErrorsDataSlotName);
			Dictionary<int, int> suppressedErrors = (Dictionary<int, int>)Thread.GetData(slot);
			if (suppressedErrors == null)
			{
				suppressedErrors = new Dictionary<int, int>();
			}

			foreach (Int32 omigaError in omigaErrors)
			{
				int referenceCount = 0;
				if (suppressedErrors.ContainsKey(omigaError))
				{
					referenceCount = suppressedErrors[omigaError];
				}
				suppressedErrors[omigaError] = ++referenceCount;
				Thread.SetData(slot, suppressedErrors);
			}
		}

		private static void RemoveSuppressedErrors(List<int> omigaErrors)
		{
			LocalDataStoreSlot slot = Thread.GetNamedDataSlot(_suppressedErrorsDataSlotName);
			Dictionary<int, int> suppressedErrors = (Dictionary<int, int>)Thread.GetData(slot);
			if (suppressedErrors != null)
			{
				foreach (int omigaError in omigaErrors)
				{
					int referenceCount = 0;
					if (suppressedErrors.ContainsKey(omigaError))
					{
						referenceCount = suppressedErrors[omigaError];
						referenceCount = Math.Max(0, referenceCount - 1);
						if (referenceCount == 0)
						{
							suppressedErrors.Remove(omigaError);
						}
						else
						{
							suppressedErrors[omigaError] = referenceCount;
						}
					}
				}
			}
		}

		private static void PushResumeOnWarningElement(XmlElement xmlElement)
		{
			LocalDataStoreSlot slot = Thread.GetNamedDataSlot(_resumeOnWarningElementStackDataSlotName);
			Stack<XmlElement> elementStack = (Stack<XmlElement>)Thread.GetData(slot);
			if (elementStack == null)
			{
				elementStack = new Stack<XmlElement>();
			}
			elementStack.Push(xmlElement);
			Thread.SetData(slot, elementStack);
		}

		private static XmlElement PeekResumeOnWarningElement()
		{
			XmlElement xmlElement = null;
			LocalDataStoreSlot slot = Thread.GetNamedDataSlot(_resumeOnWarningElementStackDataSlotName);
			Stack<XmlElement> elementStack = (Stack<XmlElement>)Thread.GetData(slot);
			if (elementStack != null)
			{
				xmlElement = elementStack.Peek();
			}
			return xmlElement;
		}

		private static XmlElement PopResumeOnWarningElement()
		{
			XmlElement xmlElement = null;
			LocalDataStoreSlot slot = Thread.GetNamedDataSlot(_resumeOnWarningElementStackDataSlotName);
			Stack<XmlElement> elementStack = (Stack<XmlElement>)Thread.GetData(slot);
			if (elementStack != null)
			{
				xmlElement = elementStack.Pop();
				if (elementStack.Count == 0)
				{
					elementStack = null;
				}
				Thread.SetData(slot, elementStack);
			}
			return xmlElement;
		}

		#region IDisposable Members
		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}

		private void Dispose(bool disposing)
		{
			if (!_disposed)
			{
				if (disposing)
				{
					// Dispose managed resources.
				}

				// Dispose unmanaged resources.
				if (_omigaErrors != null)
				{
					RemoveSuppressedErrors(_omigaErrors);
				}

				if (_xmlElement != null)
				{
					PopResumeOnWarningElement();
				}
			}

			_disposed = true;
		}
		#endregion
	}
}
