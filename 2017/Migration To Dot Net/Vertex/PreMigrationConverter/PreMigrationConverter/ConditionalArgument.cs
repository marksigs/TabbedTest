using System;
using System.Collections.Generic;
using System.Text;
using PreMigrationConverter.Exceptions;
using System.IO;

namespace PreMigrationConverter
{
    enum DirectiveType
    {
        If,
        Else,
        End,
        Const,
        NotDefined
    }

    enum ConverterAction
    {
        NotDefined,
        Skip,   // leave for manual conversion
        RemoveTrueBranch,
        RemoveFalseBranch
    }

    class PreprocessorDirective
    {
        private string _line;
        private DirectiveType _directiveType;

        internal DirectiveType DirectiveType
        {
            get 
            {
                if (!_parsed)
                {
                    throw new NotParsedException();
                }
                return _directiveType;
            }
            //set { _directiveType = value; }
        }

        private bool _parsed;
        private TextWriter _log;

        public PreprocessorDirective(string line, TextWriter log)
        {
            if (line == null)
            {
                throw (new ArgumentNullException());
            }
            if (log == null)
            {
                throw (new ArgumentNullException());
            }

            _line = line;
            _log = log;
            _parsed = false;
            _converterAction = ConverterAction.NotDefined;
        }

        private ConverterAction _converterAction;
    
        public ConverterAction Action
        {
            get 
            {
                if (!_parsed)
                {
                    throw new NotParsedException();
                }

                return _converterAction; 
            }
            //set { _converterAction = value; }
        }

        public void Parse()
        {
            if (_line.StartsWith("#If", StringComparison.OrdinalIgnoreCase))
            {
                _directiveType = DirectiveType.If;

                string[] words = _line.Split(' ');

                if (String.Compare(words[1], "Not") == 0)
                {
                    _converterAction = DetermineActionFromLiteral(words[2], true);
                }
                else
                {
                    _converterAction = DetermineActionFromLiteral(words[1], false);
                }                
            }
            else
            {
                if (_line.StartsWith("#Else", StringComparison.OrdinalIgnoreCase))
                {
                    _directiveType = DirectiveType.Else;
                }
                else
                {
                    if (_line.StartsWith("#End", StringComparison.OrdinalIgnoreCase))
                    {
                        _directiveType = DirectiveType.End;
                    }
                    else
                    {
                        if (_line.StartsWith("#Const", StringComparison.OrdinalIgnoreCase))
                        {
                            _directiveType = DirectiveType.Const;
                        }
                        else
                        {
                            _log.WriteLine("Warning - unexpected directive: " + _line);
                        }
                    }
                }
            }
            
            _parsed = true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="word"></param>
        /// <param name="reversingWordFound">Indicates this follows a "Not"</param>
        /// <returns></returns>
        private ConverterAction DetermineActionFromLiteral(
            string word, 
            bool reversingWordFound)
        {
            ConverterAction action = ConverterAction.NotDefined;

            switch (word)
            {
                case "GENERIC_SQL":
                case "Win32":
                case "BMIDS":   //GHun keep BMIDS code
                    if (reversingWordFound)
                    {
                        action = ConverterAction.RemoveTrueBranch;
                    }
                    else
                    {
                        action = ConverterAction.RemoveFalseBranch;
                    }
                    break;
                case "USING_VSA":
                case "TIME_DO":
                case "TIMINGS":
                case "OMIGA4_DEBUG":
                case "IK_DEBUG":
                case "DEBUGDATA":
                case "WriteSQLToFile":
                case "IK_DEBUG_Q":
                case "WRITETABLESTOFILE":
                case "MIN_BUILD":
                case "GPSPM_ON":
                case "ASYNC_PRINTING":
                case "ASYNCDEBUG":
                case "NODOWNLOAD":
                case "DEBUG_ON":
                case "TEST":
                case "PROFILING":   //GHun
                    if (reversingWordFound)
                    {
                        action = ConverterAction.RemoveFalseBranch;
                    }
                    else
                    {
                        action = ConverterAction.RemoveTrueBranch;
                    }
                    break;
                case "omCRUD":
                case "IGNORECOMBOLOOKUPERRORS":
                case "EPSOM":
                case "FILENET":
                case "Gemini":
                    action = ConverterAction.Skip;
                    break;
                default:
                    throw new ArgumentException("Unexpected directive literal: " + word);
            }

            return action;
        }
    }
}
