using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;

namespace Logging
{
    public class RegistrationErrors
    {

        [ExcelFunction("BadTypes")]
        public static object BadTypes(string[] input)
        {
            return input;
        }

        [ExcelFunction(Name = "+=1++")]
        public static object InvalidName()
        {
            return "Invalid Name";
        }

        [ExcelFunction]
        public static object RepeatedName()
        {
            return "Repeated Name";
        }

        [ExcelFunction]
        public static object RepeatedName(string input)
        {
            return "Repeated Name (2)";
        }

    }
}
