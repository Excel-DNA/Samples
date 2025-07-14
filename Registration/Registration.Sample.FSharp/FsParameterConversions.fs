namespace Registration.Samples.FSharp

open System
open System.Collections.Generic
open System.Linq.Expressions
open Microsoft.FSharp.Core
open ExcelDna.Integration
open ExcelDna.Registration

module FsParameterConversions =

    let FsOptionalParameterConversion (paramType : Type) (paramReg : ExcelParameterRegistration) =
        if not paramType.IsGenericType || (paramType.GetGenericTypeDefinition() <> typedefof<Option<_>>) then
            null
        else
            let innerType = paramType.GetGenericArguments().[0]
            let input = Expression.Parameter(typeof<obj>)
            Expression.Lambda(
                Expression.Condition(
                    Expression.TypeIs(input, typeof<ExcelMissing>),
                        Expression.Constant(null, paramType),
                        Expression.Call(paramType, "Some", null, 
                            TypeConversion.GetConversion(input, innerType))),
                    input)

    [<ExcelFunctionProcessor>]
    let ProcessFsOptionalParameter(registrations: IEnumerable<IExcelFunctionInfo>, config: IExcelFunctionRegistrationConfiguration) : IEnumerable<IExcelFunctionInfo> =
        // The overload selection and delegate conversions performed by F# are not intuitive.
        let paramConvertConfig = ParameterConversionConfiguration()
                                    .AddParameterConversion( 
                                        Func<Type, ExcelParameterRegistration, LambdaExpression>(FsOptionalParameterConversion),
                                        null)
        (registrations |> fun fns -> ParameterConversionRegistration.ProcessParameterConversions (fns |> Seq.cast<ExcelFunctionRegistration>, paramConvertConfig)) |> Seq.cast<IExcelFunctionInfo>