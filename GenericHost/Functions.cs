using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace GenericHost
{
    public class Functions
    {
        [ExcelFunction]
        public static object TestFromGenericHost() => "Hello from Generic Host";
    }
}
