using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using PreMigrationConverter.Properties;

namespace PreMigrationConverter
{
    /// <summary>
    /// Converts all the code in a particular file
    /// </summary>
    internal class FileConverter
    {
        private enum CodeRemovalAction
        {
            DeleteUnwantedCodeLines,
            CommentUnwantedCodeLines
        }

        private enum Condition
        {
            TrueBranch,
            FalseBranch
        }

        private string _inFileName;
        private string _outFileName;
        private TextReader _tr;
        private TextWriter _tw;
        private TextWriter _log;
        CodeRemovalAction _crt;
        private bool _fileAmended;
        private bool _logVerbose;

        public bool FileAmended
        {
            get { return _fileAmended; }
            //set { _fileAmended = value; }
        }

        public FileConverter(
            string inFileName, 
            string outFileName, 
            TextWriter log) 
        {
            if (inFileName == null || outFileName == null || log == null)
            {
                throw new ArgumentNullException();
            }

            _inFileName = inFileName;
            _outFileName = outFileName;
            _log = log;
            _fileAmended = false;
            _logVerbose = Settings.Default.LogVerbose;
            if (Settings.Default.DeleteUnwantedCode)
            {
                _crt = CodeRemovalAction.DeleteUnwantedCodeLines;
            }
            else
            {
                _crt = CodeRemovalAction.CommentUnwantedCodeLines;
            }
        }

        public bool Convert()
        {
            if (_logVerbose)
            {
                _log.WriteLine("Starting to convert file " + _inFileName);
                _log.WriteLine("... to " + _outFileName);
            }

            bool success = true;
            int lineNumber = 0;
            int totalLinesRemoved = 0;
			int totalCompilerDirectivesRemoved = 0;

            //GHun Specify encoding so that ASCII values >= 128  (e.g. copyright symbol) are not corrupted
            using (_tr =  new StreamReader(_inFileName, Encoding.GetEncoding(1252)))
            {
                using (_tw = new StreamWriter(_outFileName, false, Encoding.Default))
                {
                    string line;

                    while (((line = _tr.ReadLine()) != null) && success)
                    {
                        lineNumber++;
                        try
                        {
							int compilerDirectivesRemoved = 0;
							int linesRemoved = ConvertLine(line, out compilerDirectivesRemoved);
                            totalLinesRemoved += linesRemoved;
							totalCompilerDirectivesRemoved += compilerDirectivesRemoved;
                            lineNumber += linesRemoved;
                        }
                        catch (Exception ex)
                        {
                            _log.WriteLine("Error in line " + lineNumber.ToString() + ": " +
                                ex.Message.ToString());
                            _log.WriteLine(ex.StackTrace);
                            _log.WriteLine(line);
                            success = false;
                            throw;
                        }
                    }
                    _tw.Flush();
                }
            }

            if (success)
            {
                //GHun Overwrite the input file with the output file, creating a backup first
                if (Settings.Default.OverwriteFiles)
                {
					if (!File.Exists(_inFileName + ".orig"))
					{
						File.Move(_inFileName, _inFileName + ".orig");
					}
					else
					{
						File.Delete(_inFileName);
					}
                    File.Move(_outFileName, _inFileName);
                }
                //GHun End

                if (_logVerbose)
                {
                    _log.WriteLine("Finished converting " + lineNumber.ToString() +
                        " lines in file " + _inFileName +
						"; removed " + totalLinesRemoved.ToString() + " lines and " + totalCompilerDirectivesRemoved.ToString() + " compiler directives");
                }
                else
                {
                    _log.WriteLine(_inFileName + "," + lineNumber.ToString() +
						"," + totalLinesRemoved.ToString() + "," + totalCompilerDirectivesRemoved.ToString());
                }
            }

            return success;
        }

        private int ConvertLine(string line, out int compilerDirectivesRemoved)
        {
            int linesRemoved = 0;
			compilerDirectivesRemoved = 0;

            CodeLine cl = new CodeLine(line, _log);
            cl.Parse();

			if (cl.IsDefectComment() && _crt == CodeRemovalAction.DeleteUnwantedCodeLines)
			{
				// AS 26/11/2007 Do nothing - will remove previously commented line.
			}
			else if (cl.IsRemoveTrueBranch())
			{
				compilerDirectivesRemoved++;
				string message = GenerateMessage();
				if (_crt == CodeRemovalAction.CommentUnwantedCodeLines)
				{
					_tw.WriteLine("'" + message);
				}
				if (_logVerbose)
				{
					_log.WriteLine(message + " (removing true branch)");
				}
				linesRemoved = RemoveCondition(line, Condition.TrueBranch, ref compilerDirectivesRemoved);
				_fileAmended = true;
			}
			else if (cl.IsRemoveFalseBranch())
			{
				compilerDirectivesRemoved++;
				string message = GenerateMessage();
				if (_crt == CodeRemovalAction.CommentUnwantedCodeLines)
				{
					_tw.WriteLine("'" + message);
				}
				if (_logVerbose)
				{
					_log.WriteLine(message + " (removing false branch)");
				}
				linesRemoved = RemoveCondition(line, Condition.FalseBranch, ref compilerDirectivesRemoved);
				_fileAmended = true;
			}
			else
			{
				_tw.WriteLine(line);
			}
            return linesRemoved;
        }

        private string GenerateMessage()
        {
            string message =
                Settings.Default.DefectId + " " +
                System.DateTime.UtcNow.ToShortDateString() + " " +
                "Amended code by auto update";
            return message;
        }

        private void WriteLineAsCommentOrOmit(string line)
        {
            if (_crt == CodeRemovalAction.CommentUnwantedCodeLines)
            {
				// AS 26/11/2007 Prefix all commented lines with defect id, 
				// so we can remove them automatically later.
				_tw.Write("'" + Settings.Default.DefectId + " ");
                _tw.WriteLine(line);
            }
        }

        private int RemoveCondition(
            string startLine, 
            Condition conditionBranchToRemove,
			ref int compilerDirectivesRemoved)
        {
            int linesRemoved = 0;
            bool inTrueConditionBranch = true;

			WriteLineAsCommentOrOmit(startLine);
            linesRemoved++;

            string line;
            bool directiveEndFound = false;

            while (!directiveEndFound && ((line = _tr.ReadLine()) != null)) //GHun swap conditions around to prevent line from being read and ignored
            {
                CodeLine cl = new CodeLine(line, _log);
                cl.Parse();

                if (cl.IsDirectiveIf())
                {
                    _log.WriteLine("Nested conditional found");
                    _log.WriteLine(line);
                    throw new Exceptions.ParsingErrorException();
                }

                if (cl.IsDirectiveEnd())
                {
                    WriteLineAsCommentOrOmit(line);
                    linesRemoved++;
					compilerDirectivesRemoved++;
                    directiveEndFound = true;
                }
                else
                {
                    if (cl.IsDirectiveElse())
                    {
                        inTrueConditionBranch = false;
                        WriteLineAsCommentOrOmit(line);
                        linesRemoved++;
						compilerDirectivesRemoved++;
                    }
                    else
                    {
                        linesRemoved += ProcessLineInConditionBranch(
                            conditionBranchToRemove, 
                            inTrueConditionBranch, 
                            line);
                    }
                }
            }

            if (!directiveEndFound)
            {
                _log.WriteLine("Expected end if");
                _log.Flush();
                throw new Exceptions.ParsingErrorException();
            }

            return linesRemoved;
        }

        private int ProcessLineInConditionBranch(
            Condition conditionBranchToRemove, 
            bool inTrueCondition, 
            string line)
        {
            int linesRemoved = 0;
            if (inTrueCondition)
            {
				if (conditionBranchToRemove == Condition.TrueBranch)
				{
					WriteLineAsCommentOrOmit(line);
					linesRemoved++;
				}
				else
				{
					_tw.WriteLine(line);
				}
            }
            else
            {
                if (conditionBranchToRemove == Condition.TrueBranch)
                {
					_tw.WriteLine(line);
                }
                else
                {
                    WriteLineAsCommentOrOmit(line);
                    linesRemoved++;
                }
            }
            return linesRemoved;
        }
    }
}
