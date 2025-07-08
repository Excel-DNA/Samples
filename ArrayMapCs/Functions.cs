using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using ExcelDna.Integration.XlCall;
using Microsoft.VisualBasic.FileIO;

public static class Functions
{
    [ExcelFunction(Name="ARRAY.MAP", Description="Evaluates the given function for arrays of input values. ")]
    public static object ArrayMapN(
        [ExcelArgument(Name="function", Description="The function to evaluate - either enter the name without any quotes or brackets (for .xll functions), or as a string (for VBA functions)")] object funcNameOrId,
        [ExcelArgument(Description="The input value(s) for the first argument (row, column or rectangular range) ")] object input1,
        [ExcelArgument(Description="The input value(s) for the second argument (row, column or rectangular range) ")] object input2,
        [ExcelArgument(Description="The input value(s) for the third argument (row, column or rectangular range) ")] object input3,
        [ExcelArgument(Description="The input value(s) for the next argument (row, column or rectangular range) ")] object input4,
        [ExcelArgument(Description="The input value(s) for the next argument (row, column or rectangular range) ")] object input5,
        [ExcelArgument(Description="The input value(s) for the next argument (row, column or rectangular range) ")] object input6,
        [ExcelArgument(Description="The input value(s) for the next argument (row, column or rectangular range) ")] object input7,
        [ExcelArgument(Description="The input value(s) for the next argument (row, column or rectangular range) ")] object input8,
        [ExcelArgument(Description="The input value(s) for the next argument (row, column or rectangular range) ")] object input9,
        [ExcelArgument(Description="The input value(s) for the next argument (row, column or rectangular range) ")] object input10,
        [ExcelArgument(Description="The input value(s) for the next argument (row, column or rectangular range) ")] object input11,
        [ExcelArgument(Description="The input value(s) for the next argument (row, column or rectangular range) ")] object input12,
        [ExcelArgument(Description="The input value(s) for the next argument (row, column or rectangular range) ")] object input13,
        [ExcelArgument(Description="The input value(s) for the next argument (row, column or rectangular range) ")] object input14,
        [ExcelArgument(Description="The input value(s) for the next argument (row, column or rectangular range) ")] object input15,
        [ExcelArgument(Description="The input value(s) for the next argument (row, column or rectangular range) ")] object input16)
    {
        object[] inputs = new object[16];
        inputs[0] = input1;
        inputs[1] = input2;
        inputs[2] = input3;
        inputs[3] = input4;
        inputs[4] = input5;
        inputs[5] = input6;
        inputs[6] = input7;
        inputs[7] = input8;
        inputs[8] = input9;
        inputs[9] = input10;
        inputs[10] = input11;
        inputs[11] = input12;
        inputs[12] = input13;
        inputs[13] = input14;
        inputs[14] = input15;
        inputs[15] = input16;

        int lastPresent = 0; // 0-based index of the last input that is not Missing
        for (int i = inputs.Length - 1; i >= 0; i--)
        {
            if (!(inputs[i] is ExcelMissing))
            {
                lastPresent = i;
                break;
            }
        }

        object udfIdentifier;
        if (funcNameOrId is double)
        {
            udfIdentifier = funcNameOrId;
        }
        else if (funcNameOrId is string)
        {
            object registerId = XlCall.Excel(XlCall.xlfEvaluate, funcNameOrId);
            if (registerId is double)
                udfIdentifier = registerId;
            else
                udfIdentifier = funcNameOrId;
        }
        else
        {
            return ExcelError.ExcelErrorValue;
        }

        Func<object[], object> evaluate = args =>
        {
            object[] evalInput = new object[args.Length + 1];
            evalInput[0] = udfIdentifier;
            Array.Copy(args, 0, evalInput, 1, args.Length);
            return XlCall.Excel(XlCall.xlUDF, evalInput);
        };

        var rowInputs = new List<int>();
        var colInputs = new List<int>();

        for (int i = 0; i <= lastPresent; i++)
        {
            if (inputs[i] is object[,] arr)
            {
                int rows = arr.GetLength(0);
                int cols = arr.GetLength(1);

                if (rows > 1)
                {
                    if (cols > 1)
                        return ExcelError.ExcelErrorValue;

                    colInputs.Add(i);
                }
                else if (cols > 1)
                {
                    rowInputs.Add(i);
                }
            }
        }

        int numOutRows;
        int numOutCols;

        if (colInputs.Count == 0)
        {
            numOutRows = 1;
        }
        else
        {
            var firstColInput = (object[,])inputs[colInputs[0]];
            numOutRows = firstColInput.GetLength(0);
        }

        if (rowInputs.Count == 0)
        {
            numOutCols = 1;
        }
        else
        {
            var firstRowInput = (object[,])inputs[rowInputs[0]];
            numOutCols = firstRowInput.GetLength(1);
        }

        object[,] output = new object[numOutRows, numOutCols];

        for (int i = 0; i < numOutRows; i++)
        {
            for (int j = 0; j < numOutCols; j++)
            {
                object[] args = new object[lastPresent + 1];

                for (int index = 0; index <= lastPresent; index++)
                {
                    if (rowInputs.Contains(index))
                    {
                        args[index] = ((object[,])inputs[index])[0, j];
                    }
                    else if (colInputs.Contains(index))
                    {
                        args[index] = ((object[,])inputs[index])[i, 0];
                    }
                    else
                    {
                        if (inputs[index] is object[,] single)
                            args[index] = single[0, 0];
                        else
                            args[index] = inputs[index];
                    }
                }

                output[i, j] = evaluate(args);
            }
        }

        return output;
    }

#if DEBUG
    [ExcelFunction(IsHidden=true)]
    public static object Describe1(object x)
    {
        return x.ToString();
    }

    [ExcelFunction(IsHidden=true)]
    public static object Describe2(object x, object y)
    {
        return x.ToString() + "|" + y.ToString();
    }

    [ExcelFunction(IsHidden=true)]
    public static object TestArray()
    {
        return new object[,] { { "x" } };
    }
#endif

    [ExcelFunction(Name="ARRAY.FROMFILE", Description="Reads the contents of a delimited file")]
    public static object ArrayFromFile([ExcelArgument("Full path to the file to read")] string Path,
                                       [ExcelArgument(Name="[SkipHeader]", Description="Skips the first line of the file - default False")] object skipHeader,
                                       [ExcelArgument(Name="[Delimiter]", Description="Sets the delimiter to accept - default ','")] object delimiter)
    {
        var lines = new List<string[]>();

        using (var csvParser = new TextFieldParser(Path))
        {
            if (delimiter is ExcelMissing)
                csvParser.SetDelimiters(new string[] { "," });
            else
                csvParser.SetDelimiters(new string[] { delimiter.ToString() }); // TODO: Accept multiple ?

            csvParser.CommentTokens = new string[] { "#" };
            csvParser.HasFieldsEnclosedInQuotes = true;

            if (!(skipHeader is ExcelMissing) && (Equals(skipHeader, true) || Equals(skipHeader, 1)))
            {
                csvParser.ReadLine();
            }

            while (!csvParser.EndOfData)
            {
                lines.Add(csvParser.ReadFields());
            }
        }

        if (lines.Count == 0)
            return "";

        object[,] result = new object[lines.Count, lines[0].Length];
        for (int i = 0; i < lines.Count; i++)
        {
            for (int j = 0; j < lines[0].Length; j++)
            {
                result[i, j] = lines[i][j];
            }
        }
        return result;
    }

    [ExcelFunction(Name="ARRAY.SKIPROWS", Description="Returns the remainder of an array after skipping the first n rows")]
    public static object ArraySkipRows([ExcelArgument(AllowReference = true)] object array, int rowsToSkip)
    {
        if (array is ExcelReference arrayRef)
        {
            return new ExcelReference(arrayRef.RowFirst + rowsToSkip, arrayRef.RowLast, arrayRef.ColumnFirst, arrayRef.ColumnLast, arrayRef.SheetId);
        }
        else if (array is object[,] arrayIn)
        {
            object[,] result = new object[arrayIn.GetLength(0) - rowsToSkip, arrayIn.GetLength(1)];
            for (int i = 0; i < result.GetLength(0); i++)
            {
                for (int j = 0; j < result.GetLength(1); j++)
                {
                    result[i, j] = arrayIn[i + rowsToSkip, j];
                }
            }
            return result;
        }
        else
        {
            return array;
        }
    }

    [ExcelFunction(Name="ARRAY.COLUMN", Description="Returns a specified column from an array")]
    public static object ArrayColumn([ExcelArgument(AllowReference = true)] object array, [ExcelArgument("One-based column index to select")] int ColIndex)
    {
        if (array is ExcelReference arrayRef)
        {
            return new ExcelReference(arrayRef.RowFirst, arrayRef.RowLast, arrayRef.ColumnFirst + ColIndex - 1, arrayRef.ColumnFirst + ColIndex - 1, arrayRef.SheetId);
        }
        else if (array is object[,] arrayIn)
        {
            object[,] result = new object[arrayIn.GetLength(0), 1];
            int j = ColIndex - 1;
            for (int i = 0; i < result.GetLength(0); i++)
            {
                result[i, 0] = arrayIn[i, j];
            }
            return result;
        }
        else
        {
            return array;
        }
    }

    [ExcelFunction(IsHidden = true)]
    public static object ArrayConcat(object input1, object input2, object input3, object input4)
    {
        return $"{input1} | {input2} | {input3} | {input4}";
    }
}
