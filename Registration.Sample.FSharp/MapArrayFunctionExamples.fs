namespace Registration.Samples.FSharp

open System
open System.Collections.Generic
open System.Threading
open System.Net
open Microsoft.FSharp.Control.WebExtensions
open ExcelDna.Integration
open ExcelDna.Registration
open ExcelDna.Registration.FSharp

module MapArrayFunctionExamples =

    /// Example demonstrating the use of a simple sequence as input and output.
    /// Use an array formula in Excel to capture the output.
    [<ExcelMapArrayFunction>]
    let dnaFsRemoveDuplicates (input:seq<obj>) :seq<obj> =
        Seq.distinctBy id input

    /// A record type to use as an input parameter
    type DateBidAsk = {
        Date : System.DateTime;
        Bid : double;
        Ask : double;
    }

    /// A record type to use as an output parameter
    type DateMid = {
        Date : System.DateTime;
        Mid : double;
    }

    /// Example demonstrating an F# function which uses sequences of record types for input and output.
    /// Column headers inside the worksheet are mapped to/from record property names.
    [<ExcelMapArrayFunction>]
    let dnaFsCalculateMids ([<ExcelMapPropertiesToColumnHeaders>] input:seq<DateBidAsk>) : 
        [<ExcelMapPropertiesToColumnHeaders>] seq<DateMid> = 

        let CalculateMid (input:DateBidAsk) :DateMid = 
            { Date = input.Date; Mid = (input.Bid + input.Ask)/2.}
        Seq.map CalculateMid input

