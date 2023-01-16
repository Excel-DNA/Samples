namespace Registration.Samples.FSharp

open System
open System.Linq.Expressions
open ExcelDna.Integration
open ExcelDna.Registration
open ExcelDna.Registration.FSharp

type FsAsyncAddIn () =
    interface IExcelAddIn with
        member this.AutoOpen ()  = 
            // The overload selection and delegate conversions performed by F# are not intuitive.
            let paramConvertConfig = ParameterConversionConfiguration()
                                        .AddParameterConversion( 
                                            Func<Type, ExcelParameterRegistration, LambdaExpression>(FsParameterConversions.FsOptionalParameterConversion),
                                            null)

            ExcelRegistration.GetExcelFunctions ()
            |> fun fns -> ParameterConversionRegistration.ProcessParameterConversions (fns, paramConvertConfig)
            |> FsAsyncRegistration.ProcessFsAsyncRegistrations
            |> AsyncRegistration.ProcessAsyncRegistrations
            |> MapArrayFunctionRegistration.ProcessMapArrayFunctions
            |> ExcelRegistration.RegisterFunctions
        
        member this.AutoClose () = ()
