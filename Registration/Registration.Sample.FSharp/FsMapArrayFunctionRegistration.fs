namespace Registration.Samples.FSharp

open ExcelDna.Integration
open System.Collections.Generic

module FsMapArrayFunctionRegistration =
    [<ExcelFunctionProcessor>]
    let ProcessMapArrayFunctions(registrations: IEnumerable<IExcelFunctionInfo>, config: IExcelFunctionRegistrationConfiguration) : IEnumerable<IExcelFunctionInfo> =
        ExcelDna.Registration.MapArrayFunctionRegistration.ProcessMapArrayFunctions(registrations, config)

