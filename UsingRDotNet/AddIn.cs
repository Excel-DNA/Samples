using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using ExcelDna.Integration;
using ExcelDna.Logging;
using RDotNet;

namespace UsingRDotNet
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            MyFunctions.InitializeRDotNet();
        }

        public void AutoClose()
        {
        }
    }

    public static class MyFunctions
    {
        static REngine _engine;
        internal static void InitializeRDotNet()
        {
            try
            {
                REngine.SetEnvironmentVariables();
                _engine = REngine.GetInstance();
                _engine.Initialize();
            }
            catch (Exception ex)
            {
                LogDisplay.WriteLine("Error initializing RDotNet: " + ex.Message);
            }
        }

        public static double[] MyRnorm(int number)
        {
            return (_engine.Evaluate("rnorm(" + number + ")").AsNumeric().ToArray<double>());
        }

        public static object TestRDotNet()
        {
            // .NET Framework array to R vector.
            NumericVector group1 = _engine.CreateNumericVector(new double[] { 30.02, 29.99, 30.11, 29.97, 30.01, 29.99 });
            _engine.SetSymbol("group1", group1);
            // Direct parsing from R script.
            NumericVector group2 = _engine.Evaluate("group2 <- c(29.89, 29.93, 29.72, 29.98, 30.02, 29.98)").AsNumeric();

            // Test difference of mean and get the P-value.
            GenericVector testResult = _engine.Evaluate("t.test(group1, group2)").AsList();
            double p = testResult["p.value"].AsNumeric().First();

            return string.Format("Group1: [{0}], Group2: [{1}], P-value = {2:0.000}",  string.Join(", ", group1), string.Join(", ", group2), p);
        }
    }
}
