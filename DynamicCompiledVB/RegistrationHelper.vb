Imports System.Reflection
Imports ExcelDna.Registration
Imports ExcelDna.Registration.VisualBasic

Friend Module RegistrationHelper

    Sub PerformRegistration(assemblies As IEnumerable(Of Assembly))

        Dim conversionConfig As ParameterConversionConfiguration
        conversionConfig = New ParameterConversionConfiguration() _
                            .AddParameterConversion(ParameterConversions.GetOptionalConversion(treatEmptyAsMissing:=False)) _
                            .AddParameterConversion(AddressOf RangeParameterConversion.ParameterConversion, Nothing)

        GetAllPublicSharedFunctions(assemblies) _
        .ProcessParamsRegistrations() _
        .UpdateRegistrationsForRangeParameters() _
        .ProcessParameterConversions(conversionConfig) _
        .RegisterFunctions()

        GetAllPublicSharedSubs(assemblies).RegisterCommands()
    End Sub

    ' Gets the Public Shared methods that don't return Void
    Private Function GetAllPublicSharedFunctions(assemblies As IEnumerable(Of Assembly)) As IEnumerable(Of ExcelFunctionRegistration)
        Return From ass In assemblies
               From typ In ass.GetTypes()
               Where Not typ.FullName.Contains(".My.") AndAlso typ.IsPublic
               From mi In typ.GetMethods(BindingFlags.Public Or BindingFlags.Static)
               Where Not mi.ReturnType = GetType(Void)
               Select New ExcelFunctionRegistration(mi)
    End Function

    ' Gets the Public Shared methods that return Void
    Private Function GetAllPublicSharedSubs(assemblies As IEnumerable(Of Assembly)) As IEnumerable(Of ExcelCommandRegistration)
        Return From ass In assemblies
               From typ In ass.GetTypes()
               Where Not typ.FullName.Contains(".My.") AndAlso typ.IsPublic
               From mi In typ.GetMethods(BindingFlags.Public Or BindingFlags.Static)
               Where mi.ReturnType = GetType(Void)
               Select New ExcelCommandRegistration(mi)
    End Function

End Module

