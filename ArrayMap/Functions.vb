Imports ExcelDna.Integration
Imports ExcelDna.Integration.XlCall

Public Module Functions

    <ExcelFunction(Name:="ARRAY.MAP", Description:="Evaluates the function for every value in the input array, returning an array that has the same size as the input.")>
    Function ArrayMap(
                     <ExcelArgument(Description:="The function to evaluate - either enter the name without any quotes or brackets (for .xll functions), or as a string (for VBA functions)")>
                     [function] As Object,
                     <ExcelArgument(Description:="The array of input values (row, column or rectangular range) ")>
                     input As Object)

        Dim evaluate As Func(Of Object, Object)

        If TypeOf [function] Is Double Then
            evaluate = Function(x) Excel(xlUDF, [function], x)
        ElseIf TypeOf [function] Is String Then
            ' First try to get the RegisterId, if it's an .xll UDF
            Dim registerId As Object
            registerId = Excel(xlfEvaluate, [function])
            If TypeOf registerId Is Double Then
                evaluate = Function(x) Excel(xlUDF, registerId, x)
            Else
                ' Just call as string, hoping it's a valid VBA function
                evaluate = Function(x) Excel(xlUDF, [function], x)
            End If
        Else
            Return ExcelError.ExcelErrorValue
        End If

        Return ArrayEvaluate(evaluate, input)
    End Function

    <ExcelFunction(Name:="ARRAY.MAP2", Description:="Evaluates the two-argument function for every value in the first and second inputs. " &
                   "Takes a single value and any rectangle, or one row and one column, or one column and one row.")>
    Function ArrayMap2(
                     <ExcelArgument(Description:="The function to evaluate - either enter the name without any quotes or brackets (for .xll functions), or as a string (for VBA functions)")>
                     [function] As Object,
                     <ExcelArgument(Description:="The input value(s) for the first argument (row, column or rectangular range) ")>
                     input1 As Object,
                     <ExcelArgument(Description:="The input value(s) for the second argument (row, column or rectangular range) ")>
                     input2 As Object)

        Dim evaluate As Func(Of Object, Object, Object)

        If TypeOf [function] Is Double Then
            evaluate = Function(x, y) Excel(xlUDF, [function], x, y)
        ElseIf TypeOf [function] Is String Then
            ' First try to get the RegisterId, if it's an .xll UDF
            Dim registerId As Object
            registerId = Excel(xlfEvaluate, [function])
            If TypeOf registerId Is Double Then
                evaluate = Function(x, y) Excel(xlUDF, registerId, x, y)
            Else
                ' Just call as string, hoping it's a valid VBA function
                evaluate = Function(x, y) Excel(xlUDF, [function], x, y)
            End If
        Else
            Return ExcelError.ExcelErrorValue
        End If

        If Not TypeOf input1 Is Object(,) Then
            Dim evaluate1 = Function(x) evaluate(input1, x)
            Return ArrayEvaluate(evaluate1, input2)
        ElseIf Not TypeOf input2 Is Object(,) Then
            Dim evaluate1 = Function(x) evaluate(x, input2)
            Return ArrayEvaluate(evaluate1, input1)
        End If

        ' Now we know both input1 and input2 are arrays
        ' We assume they are 1D, else error
        If input1.GetLength(0) > 1 Then

            ' Lots of rows in input1, we'll take it's first column only, and take the columns input1
            Dim output(input1.GetLength(0) - 1, input2.GetLength(1) - 1) As Object

            For i As Integer = 0 To input1.GetLength(0) - 1
                For j As Integer = 0 To input2.GetLength(1) - 1
                    output(i, j) = evaluate(input1(i, 0), input2(0, j))
                Next
            Next
            Return output
        Else

            ' Single row in input1, we'll take it's columns, and take the rows from input2
            Dim output(input2.GetLength(0) - 1, input1.GetLength(1) - 1) As Object

            For i As Integer = 0 To input2.GetLength(0) - 1
                For j As Integer = 0 To input1.GetLength(1) - 1
                    output(i, j) = evaluate(input1(0, j), input2(i, 0))
                Next
            Next
            Return output
        End If

    End Function

    Private Function ArrayEvaluate(evaluate As Func(Of Object, Object), input As Object) As Object
        If TypeOf input Is Object(,) Then
            Dim output(input.GetLength(0) - 1, input.GetLength(1) - 1) As Object

            For i As Integer = 0 To input.GetLength(0) - 1
                For j As Integer = 0 To input.GetLength(1) - 1
                    output(i, j) = evaluate(input(i, j))
                Next
            Next
            Return output
        Else
            Return evaluate(input)
        End If
    End Function

    <ExcelFunction(IsHidden:=True)>
    Function Describe1(x)
        Return x.ToString()
    End Function

    <ExcelFunction(IsHidden:=True)>
    Function Describe2(x, y)
        Return x.ToString() & "|" & y.ToString()
    End Function

End Module
