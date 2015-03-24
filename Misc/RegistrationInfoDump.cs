using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using System.Diagnostics;

public static class RegistrationInfoDump
{

    // This macro retrieves and dumpt to a new sheet (in the ActiveWorkbook) a list of all the loaded add-ins, 
    // as well as the full registration info for loaded Excel-DNA add-ins.
    [ExcelCommand(MenuName = "Excel-DNA", MenuText = "Registration Info Dump")]
    public static void RegistrationInfo()
    {
        try 
        {
            Excel.Application Application = ExcelDnaUtil.Application as Excel.Application;
            List<string> addinPaths = new List<string>();

            if (ExcelDnaUtil.ExcelVersion >= 14.0) 
            {
                foreach (dynamic addIn in Application.AddIns2)
                {
                    if (addIn.IsOpen)
                    {
                        addinPaths.Add(addIn.FullName);
                    }
                }
            } 
            else 
            {
                HashSet<string> allPaths = new HashSet<string>();
                dynamic funcInfos = Application.RegisteredFunctions;
                if ((funcInfos != null)) 
                {
                    for (int i = funcInfos.GetLowerBound(0); i <= funcInfos.GetUpperBound(0); i++) 
                    {
                        allPaths.Add(funcInfos[i, 1]);
                    }
                }
                addinPaths.AddRange(allPaths);
            }

            dynamic wb = Application.Workbooks.Add();
            Excel.Worksheet shIndex = wb.Sheets(1);
            shIndex.Name = "Add-Ins";
            shIndex.Cells[1, 1] = "Add-In Path";
            shIndex.Cells[1, 2] = "Registration Info?";

            int row = 2;
            foreach (string path in addinPaths) 
            {
                shIndex.Cells[row, 1] = path;

                // Try to read RegistrationInfo
                dynamic result = ExcelIntegration.GetRegistrationInfo(path, 0);
                if (result.Equals(ExcelError.ExcelErrorNA)) 
                {
                    shIndex.Cells[row, 2] = false;
                } 
                else 
                {
                    shIndex.Cells[row, 2] = true;

                    // Dump the result to a new sheet
                    Excel.Worksheet shInfo = wb.Sheets.Add(After: wb.Sheets(wb.Sheets.Count));
                    shInfo.Name = System.IO.Path.GetFileName(path);

                    // C API via ExcelReference would work well here
                    dynamic refInfo = new ExcelReference(0, result.GetUpperBound(0), 0, 254, shInfo.Name);
                    refInfo.SetValue(result);
                }
                row = row + 1;
            }
            shIndex.Activate();

        } 
        catch (Exception ex) 
        {
            Debug.Print(ex.ToString());
        }
    }
}
