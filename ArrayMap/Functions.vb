Imports ExcelDna.Integration
Imports ExcelDna.Integration.XlCall
Imports Microsoft.VisualBasic.FileIO

Public Module Functions
    <ExcelFunction(Name:="ARRAY.MAP", Description:="Evaluates the given function for arrays of input values. ")>
    Function ArrayMapN(
                     <ExcelArgument(Name:="function", Description:="The function to evaluate - either enter the name without any quotes or brackets (for .xll functions), or as a string (for VBA functions)")>
                     funcNameOrId As Object,
                     <ExcelArgument(Description:="The input value(s) for the first argument (row, column or rectangular range) ")>
                     input1 As Object,
                     <ExcelArgument(Description:="The input value(s) for the second argument (row, column or rectangular range) ")>
                     input2 As Object,
                     <ExcelArgument(Description:="The input value(s) for the third argument (row, column or rectangular range) ")>
                     input3 As Object,
                     <ExcelArgument(Description:="The input value(s) for the next argument (row, column or rectangular range) ")>
                     input4 As Object,
                     <ExcelArgument(Description:="The input value(s) for the next argument (row, column or rectangular range) ")>
                     input5 As Object,
                     <ExcelArgument(Description:="The input value(s) for the next argument (row, column or rectangular range) ")>
                     input6 As Object,
                     <ExcelArgument(Description:="The input value(s) for the next argument (row, column or rectangular range) ")>
                     input7 As Object,
                     <ExcelArgument(Description:="The input value(s) for the next argument (row, column or rectangular range) ")>
                     input8 As Object,
                     <ExcelArgument(Description:="The input value(s) for the next argument (row, column or rectangular range) ")>
                     input9 As Object,
                     <ExcelArgument(Description:="The input value(s) for the next argument (row, column or rectangular range) ")>
                     input10 As Object,
                     <ExcelArgument(Description:="The input value(s) for the next argument (row, column or rectangular range) ")>
                     input11 As Object,
                     <ExcelArgument(Description:="The input value(s) for the next argument (row, column or rectangular range) ")>
                     input12 As Object,
                     <ExcelArgument(Description:="The input value(s) for the next argument (row, column or rectangular range) ")>
                     input13 As Object,
                     <ExcelArgument(Description:="The input value(s) for the next argument (row, column or rectangular range) ")>
                     input14 As Object,
                     <ExcelArgument(Description:="The input value(s) for the next argument (row, column or rectangular range) ")>
                     input15 As Object,
                     <ExcelArgument(Description:="The input value(s) for the next argument (row, column or rectangular range) ")>
                     input16 As Object)


        Dim inputs(15) As Object

        inputs(0) = input1
        inputs(1) = input2
        inputs(2) = input3
        inputs(3) = input4
        inputs(4) = input5
        inputs(5) = input6
        inputs(6) = input7
        inputs(7) = input8
        inputs(8) = input9
        inputs(9) = input10
        inputs(10) = input11
        inputs(11) = input12
        inputs(12) = input13
        inputs(13) = input14
        inputs(14) = input15
        inputs(15) = input16

        Dim evaluate As Func(Of Object(), Object)

        Dim lastPresent As Integer ' 0-based index of the last input that is not Missing
        For i = inputs.Length - 1 To 0 Step -1
            If TypeOf inputs(i) IsNot ExcelMissing Then
                lastPresent = i
                Exit For
            End If
        Next

        Dim udfIdentifier ' Either a registerId or a string name

        If TypeOf funcNameOrId Is Double Then
            udfIdentifier = funcNameOrId
        ElseIf TypeOf funcNameOrId Is String Then
            ' First try to get the RegisterId, if it's an .xll UDF
            Dim registerId As Object
            registerId = Excel(xlfEvaluate, funcNameOrId)
            If TypeOf registerId Is Double Then
                udfIdentifier = registerId
            Else
                udfIdentifier = funcNameOrId
            End If
        Else
            ' Something we don't understand
            Return ExcelError.ExcelErrorValue
        End If

        evaluate = Function(args)
                       Dim evalInput(args.Length) As Object
                       evalInput(0) = udfIdentifier
                       Array.Copy(args, 0, evalInput, 1, args.Length)
                       Return Excel(xlUDF, evalInput)
                   End Function

        ' An input argument might appear in both of these collections, if it is a non-skinny rectangle
        Dim rowInputs As New List(Of Integer)
        Dim colInputs As New List(Of Integer)

        For i As Integer = 0 To lastPresent
            If TypeOf inputs(i) Is Object(,) Then

                Dim rows = inputs(i).GetLength(0)
                Dim cols = inputs(i).GetLength(1)

                If rows > 1 Then
                    If cols > 1 Then
                        Return ExcelError.ExcelErrorValue
                    End If

                    colInputs.Add(i)
                ElseIf cols > 1 Then
                    rowInputs.Add(i)
                End If
            End If
        Next

        Dim numOutRows As Integer
        Dim numOutCols As Integer

        If colInputs.Count = 0 Then
            numOutRows = 1
        Else
            ' TODO: Check that all of the column inputs have the same length
            Dim firstColInput As Object(,) = inputs(colInputs(0))
            numOutRows = firstColInput.GetLength(0)
        End If

        If rowInputs.Count = 0 Then
            numOutCols = 1
        Else
            ' TODO: Check
            Dim firstRowInput As Object(,) = inputs(rowInputs(0))
            numOutCols = firstRowInput.GetLength(1)
        End If

        Dim output(numOutRows - 1, numOutCols - 1) As Object

        For i = 0 To numOutRows - 1
            For j As Integer = 0 To numOutCols - 1
                Dim args(lastPresent) As Object

                For index = 0 To lastPresent  ' Do this stuff for each arg index
                    If rowInputs.Contains(index) Then
                        ' inputs(index) is a row
                        args(index) = inputs(index)(0, j)
                    ElseIf colInputs.Contains(index) Then
                        ' inputs(index) is a column
                        args(index) = inputs(index)(i, 0)
                    Else
                        ' input might still be a 1x1 array, which we want to dereference
                        If TypeOf inputs(index) Is Object(,) Then
                            args(index) = inputs(index)(0, 0)
                        Else
                            args(index) = inputs(index)
                        End If
                    End If
                Next

                output(i, j) = evaluate(args)
            Next
        Next

        Return output

    End Function

#If DEBUG Then
    <ExcelFunction(IsHidden:=True)>
    Function Describe1(x)
        Return x.ToString()
    End Function

    <ExcelFunction(IsHidden:=True)>
    Function Describe2(x, y)
        Return x.ToString() & "|" & y.ToString()
    End Function

    <ExcelFunction(IsHidden:=True)>
    Function TestArray()
        Return New Object(,) {{"x"}}
    End Function
#End If


    <ExcelFunction(Name:="ARRAY.FROMFILE", Description:="Reads the contents of a delimited file")>
    Function ArrayFromFile(<ExcelArgument("Full path to the file to read")> Path As String,
                           <ExcelArgument(Name:="[SkipHeader]", Description:="Skips the first line of the file - default False")> skipHeader As Object,
                           <ExcelArgument(Name:="[Delimiter]", Description:="Sets the delimiter to accept - default ','")> delimiter As Object)

        Dim lines As New List(Of String())

        Using csvParser As New TextFieldParser(Path)


            If TypeOf delimiter Is ExcelMissing Then
                csvParser.SetDelimiters(New String() {","})
            Else
                csvParser.SetDelimiters(New String() {delimiter})   ' TODO: Accept multiple ?
            End If

            csvParser.CommentTokens = New String() {"#"}
            csvParser.HasFieldsEnclosedInQuotes = True

            If Not TypeOf skipHeader Is ExcelMissing AndAlso (skipHeader = True OrElse skipHeader = 1) Then
                csvParser.ReadLine()
            End If

            Do While csvParser.EndOfData = False
                lines.Add(csvParser.ReadFields())
            Loop
        End Using

        If lines.Count = 0 Then
            Return ""
        End If

        Dim result(lines.Count - 1, lines(0).Length - 1) As Object
        For i As Integer = 0 To lines.Count - 1
            For j As Integer = 0 To lines(0).Length - 1
                result(i, j) = lines(i)(j)
            Next j
        Next i
        Return result

    End Function

    <ExcelFunction(Name:="ARRAY.SKIPROWS", Description:="Returns the remainder of an array after skipping the first n rows")>
    Function ArraySkipRows(<ExcelArgument(AllowReference:=True)> array As Object, rowsToSkip As Integer)
        If TypeOf array Is ExcelReference Then
            Dim arrayRef As ExcelReference = array
            Return New ExcelReference(arrayRef.RowFirst + rowsToSkip, arrayRef.RowLast, arrayRef.ColumnFirst, arrayRef.ColumnLast, arrayRef.SheetId)
        ElseIf TypeOf array Is Object(,) Then
            Dim arrayIn As Object(,) = array
            Dim result(array.GetLength(0) - rowsToSkip - 1, array.GetLength(1) - 1) As Object
            For i As Integer = 0 To result.GetLength(0) - rowsToSkip - 1
                For j As Integer = 0 To result.GetLength(1) - 1
                    result(i, j) = arrayIn(i + rowsToSkip, j)
                Next j
            Next i
            Return result
        Else
            Return array
        End If
    End Function

    <ExcelFunction(Name:="ARRAY.COLUMN", Description:="Returns a specified column from an array")>
    Function ArrayColumn(<ExcelArgument(AllowReference:=True)> array As Object, <ExcelArgument("One-based column index to select")> ColIndex As Integer)
        If TypeOf array Is ExcelReference Then
            Dim arrayRef As ExcelReference = array
            Return New ExcelReference(arrayRef.RowFirst, arrayRef.RowLast, arrayRef.ColumnFirst + ColIndex - 1, arrayRef.ColumnFirst + ColIndex - 1, arrayRef.SheetId)
        ElseIf TypeOf array Is Object(,) Then
            Dim arrayIn As Object(,) = array
            Dim result(array.GetLength(0) - 1, 1) As Object
            Dim j As Integer = ColIndex - 1
            For i As Integer = 0 To result.GetLength(0) - 1
                result(i, 0) = arrayIn(i, j)
            Next i
            Return result
        Else
            Return array
        End If
    End Function

    <ExcelFunction(IsHidden:=True)>
    Function ArrayConcat(input1 As Object, input2 As Object, input3 As Object, input4 As Object)
        Return $"{input1} | {input2} | {input3} | {input4}"
    End Function

End Module
