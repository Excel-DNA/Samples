using ExcelDna.Integration;

namespace Registration.Sample
{
    public class Conversions
    {
        [ExcelParameterConversion]
        public static TestType1 ToTestType1(string value)
        {
            return new TestType1(value);
        }

        [ExcelParameterConversion]
        public static TestType2 TestType2FromTestType1(TestType1 value)
        {
            return new TestType2(value);
        }
    }
}
