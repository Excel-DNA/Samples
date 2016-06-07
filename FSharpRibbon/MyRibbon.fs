module MyRibbon

open System.Windows.Forms
open System.Runtime.InteropServices
open Microsoft.Office.Interop.Excel
open ExcelDna.Integration
open ExcelDna.Integration.CustomUI

// This defines a regular Excel macro (in Excel you can press Alt + F8, type in the name "showMessage", then click the Run button).
// For the ribbon, it will be run through the ExcelRibbon.RunTagMacro(...) helper, which run whatever macro is specified in the button tag attribute
// One advantage is that you can 
[<ExcelCommand>]
let showMessage () =
    XlCall.Excel(XlCall.xlcAlert, "Hello from a macro!") 
    |> ignore


// This type defines the ribbon interface. It is a public class that derives from ExcelRibbon
type public MyRibbon() =
    inherit ExcelRibbon()

    // The ribbon xml definition could also be placed in the .dna file
    // Remember to switch on the ExcelOption "Show add-in user interface errors" option (under the Advanced tab under General)
    override this.GetCustomUI(ribbonId) = 
        @"<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' >
          <ribbon>
            <tabs>
              <tab id='CustomTab' label='My F# Tab'>
                <group id='SampleGroup' label='My Sample Group'>
                  <button id='Button1' label='Run a macro' onAction='RunTagMacro' tag='showMessage' />
                  <button id='Button2' label='Run a class member' onAction='OnButtonPressed'/>
                  <button id='Button3' label='Dump the Excel Version to cell A1' onAction='OnDumpData'/>
                </group >
              </tab>
            </tabs>
          </ribbon>
        </customUI>"

    member this.OnButtonPressed (control:IRibbonControl) =
        MessageBox.Show "Hello from F#!" 
        |> ignore

    member this.OnDumpData (control:IRibbonControl) =
        let app = ExcelDnaUtil.Application :?> Application
        let cellA1 = app.Range("A1")
        cellA1.Value2 <- app.Version
        // could also replace the last line with
        //     cellA1.Value(XlRangeValueDataType.xlRangeValueDefault) <- app.Version 
        // but Range.Value is an indexer property, so it's a bit inconvenient
