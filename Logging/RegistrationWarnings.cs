using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;

namespace Logging
{
    public class RegistrationWarnings
    {
        public static object LongArgumentNames(
            [ExcelArgument(Name = "ThisIsAVeryVeryLongName0123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789")] object arg0,
            [ExcelArgument(Name = "ThisIsAVeryVeryLongName1")] object arg1,
            [ExcelArgument(Name = "ThisIsAVeryVeryLongName2")] object arg2,
            [ExcelArgument(Name = "ThisIsAVeryVeryLongName3")] object arg3,
            [ExcelArgument(Name = "ThisIsAVeryVeryLongName4")] object arg4,
            [ExcelArgument(Name = "ThisIsAVeryVeryLongName5")] object arg5,
            [ExcelArgument(Name = "ThisIsAVeryVeryLongName6")] object arg6,
            [ExcelArgument(Name = "ThisIsAVeryVeryLongName7")] object arg7,
            [ExcelArgument(Name = "ThisIsAVeryVeryLongName8")] object arg8,
            [ExcelArgument(Name = "ThisIsAVeryVeryLongName9")] object arg9,
            [ExcelArgument(Name = "ThisIsAVeryVeryLongNameA")] object argA,
            [ExcelArgument(Name = "ThisIsAVeryVeryLongNameB")] object argB
            )
        {
            return "Hello from the long func";
        }

        public static object LongFinalArgumentNames(
            [ExcelArgument(Name = "ThisIsAVeryVeryLongName0")] object arg0,
            [ExcelArgument(Name = "ThisIsAVeryVeryLongName1")] object arg1,
            [ExcelArgument(Name = "ThisIsAVeryVeryLongName2")] object arg2,
            [ExcelArgument(Name = "ThisIsAVeryVeryLongName3")] object arg3,
            [ExcelArgument(Name = "ThisIsAVeryVeryLongName4")] object arg4,
            [ExcelArgument(Name = "ThisIsAVeryVeryLongName5")] object arg5,
            [ExcelArgument(Name = "ThisIsAVeryVeryLongName6")] object arg6,
            [ExcelArgument(Name = "ThisIsAVeryVeryLongName7")] object arg7,
            [ExcelArgument(Name = "ThisIsAVeryVeryLongName8")] object arg8,
            [ExcelArgument(Name = "ThisIsAVeryVeryLongName9")] object arg9,
            [ExcelArgument(Name = "ThisIsAVeryVeryLongNameA")] object argA,
            [ExcelArgument(Name = "ThisIsAVeryVeryLongNameB")] object argB
            )
        {
            return "Hello from the long func";
        }

        public static object BadTypesUnmarked(string[] input)
        {
            return input;
        }


        public static object LongName123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890()
        {
            return 1;
        }

        [ExcelFunction(Category = "CAT123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890")]
        public static object LongCategoryName()
        {
            return "LongCategoryName";
        }

        [ExcelFunction(Description = "DESC123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890")]
        public static object LongDescription()
        {
            return "LongDescription";
        }

        [ExcelFunction(HelpTopic = "DESC123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890")]
        public static object LongHelpTopic()
        {
            return "LongHelpTopic";
        }

        // Info level event - not registered due to explciit...
        [ExcelFunction(ExplicitRegistration = true)]
        public static object ExplicitRegistration()
        {
            return "Explicit Registration";
        }

    }
}
