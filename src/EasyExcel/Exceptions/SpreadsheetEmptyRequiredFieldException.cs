using System;

namespace EasyExcel.Exceptions
{
    public class SpreadsheetEmptyRequiredFieldException : Exception
    {
        public SpreadsheetEmptyRequiredFieldException()
        {
        }

        public SpreadsheetEmptyRequiredFieldException(string message)
            : base(message)
        {
        }

        public SpreadsheetEmptyRequiredFieldException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}
