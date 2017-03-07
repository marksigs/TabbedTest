using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using PreMigrationConverter.Exceptions;
using PreMigrationConverter.Properties;

namespace PreMigrationConverter
{
    class CodeLine
    {
        private string _line;
        private bool _parsed;
        private TextWriter _log;

        /// <summary>
        /// line is a #If, #Else or #End
        /// </summary>
        private bool _isDirective;

		/// <summary>
		/// Line is a defect comment added by a previous run.
		/// </summary>
		private bool _isDefectComment;

        private PreprocessorDirective _directive;

        public CodeLine(string line, TextWriter log)
        {
            if (line == null || log == null)
            {
                throw new ArgumentNullException();
            }
            _line = line;
            _log = log;
            _parsed = false;
        }

        /// <summary>
        /// Parses line of code into 
        /// </summary>
        public void Parse()
        {
			if (_line.StartsWith("'" + Settings.Default.DefectId))
			{
				_isDefectComment = true;
			}
			else
			{
				int numChars = _line.Length;
				int thisChar = 0;
				bool stop = false;

				while ((thisChar < numChars)
					&& !stop)
				{
					switch (_line[thisChar])
					{
						case '\'':
							// line is a comment
							stop = true;
							break;
						case ' ':
							// do nothing
							break;
						case '#':
							_isDirective = true;
							_directive = new PreprocessorDirective(
								_line.Substring(thisChar),
								_log);
							_directive.Parse();
							stop = true;
							break;
						default:
							// '#' must be the first non-space char on a line 
							stop = true;
							break;
					}
					thisChar++;
				}
			}

            _parsed = true;
        }

        /// <summary>
        /// Check if line is a #If, #Else or #End
        /// </summary>
        /// <returns></returns>
        private bool IsDirective()
        {
            if (!_parsed)
            {
                throw new NotParsedException();
            }

            return _isDirective;
        }

		public bool IsDefectComment()
		{
			if (!_parsed)
			{
				throw new NotParsedException();
			}

			return _isDefectComment;
		}

		
		/// <summary>
        /// Check if line is a #If
        /// </summary>
        /// <returns></returns>
        public bool IsDirectiveIf()
        {
            if (!_parsed)
            {
                throw new NotParsedException();
            }

            bool value = false;

            if (_isDirective)
            {
                value = (_directive.DirectiveType == DirectiveType.If);
            }

            return value;
        }

        /// <summary>
        /// Check if line is a #Else
        /// </summary>
        /// <returns></returns>
        public bool IsDirectiveElse()
        {
            if (!_parsed)
            {
                throw new NotParsedException();
            }

            bool value = false;

            if (_isDirective)
            {
                value = (_directive.DirectiveType == DirectiveType.Else);
            }

            return value;
        }

        /// <summary>
        /// Check if line is a #End
        /// </summary>
        /// <returns></returns>
        public bool IsDirectiveEnd()
        {
            if (!_parsed)
            {
                throw new NotParsedException();
            }

            bool value = false;

            if (_isDirective)
            {
                value = (_directive.DirectiveType == DirectiveType.End);
            }

            return value;
        }

        public bool IsRemoveTrueBranch()
        {
            if (!_parsed)
            {
                throw new NotParsedException();
            }

            bool value = false;

            if (IsDirectiveIf())
            {
                if (_directive.Action == ConverterAction.RemoveTrueBranch)
                {
                    value = true;
                }
            }

            return value; 
        }

        public bool IsRemoveFalseBranch()
        {
            if (!_parsed)
            {
                throw new NotParsedException();
            }

            bool value = false;

            if (IsDirectiveIf())
            {
                if (_directive.Action == ConverterAction.RemoveFalseBranch)
                {
                    value = true;
                }
            }

            return value;
        }
    }
}
