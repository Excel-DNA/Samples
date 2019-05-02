Imports ExcelDna.Integration
Imports ExcelDna.IntelliSense
Imports System.CodeDom.Compiler
Imports System.Reflection

Namespace TheAddin
    Public Class MyAddIn
        Implements IExcelAddIn

        Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
            IntelliSenseServer.Install()

            ' Get the default list of exported assemblies
            Dim assemblies As New List(Of Assembly)
            assemblies.AddRange(ExcelIntegration.GetExportedAssemblies())

            ' Compile and add the dyanmic assembly to the list
            Dim dynamicAssembly As Assembly = CreateFunctionAssembly()
            If Not dynamicAssembly Is Nothing Then
                assemblies.Add(dynamicAssembly)
            End If

            ' Use our own registration helper to do the registration
            RegistrationHelper.PerformRegistration(assemblies)
        End Sub

        Public Sub AutoClose() Implements IExcelAddIn.AutoClose
            IntelliSenseServer.Uninstall()
        End Sub

        Private Function CreateFunctionAssembly() As Assembly
            Dim dq As String = """"
            Dim c As VBCodeProvider = New VBCodeProvider
            Dim cp As CompilerParameters = New CompilerParameters

            Dim code As String = "Namespace NewFunctions
                                    Public Module MyNewUDF
                                        Function GETNAME()
                                            Return " + dq + "My name is Jeff" + dq + "
                                        End Function
                                        Function GETAGE()
                                            Return " + dq + "My age is 23" + dq + "
                                        End Function
                                    End Module
                                End Namespace"

            Dim results As CompilerResults = c.CompileAssemblyFromSource(cp, code)

            If results.Errors.HasErrors Then
                For Each er As CompilerError In results.Errors
                    MsgBox(er.ErrorNumber + ". " + er.ErrorText)
                Next
                Return Nothing
            End If

            Return results.CompiledAssembly

        End Function
    End Class
End Namespace

Namespace OldFunctions
    Public Module MyOldUDF
        Function GETHEIGHT() As String
            Return "Height is 170cm"
        End Function


        Function GETWEIGHT() As String
            Return "Weight is 70KG"
        End Function
    End Module
End Namespace
