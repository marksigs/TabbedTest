using System;
using System.Collections.Generic;
using System.Text;

namespace PreMigrationConverter.Exceptions
{
    /// <summary>
    /// Attempting to check value before line has been parsed
    /// </summary>
    class NotParsedException : Exception
    {
    }

    /// <summary>
    /// Parsing has gone wrong
    /// </summary>
    class ParsingErrorException : Exception
    {
    }

}
