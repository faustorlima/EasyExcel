using System;

namespace EasyExcel.Exceptions
{
    public class SpreadsheetValueConversionException : Exception
    {
        public SpreadsheetValueConversionException()
        {
        }

        public SpreadsheetValueConversionException(string message)
            : base(message)
        {
        }

        public SpreadsheetValueConversionException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}
