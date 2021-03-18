# RibbonVB

This sample shows some dynamic Ribbon techniques from my [DB-Addin](https://github.com/rkapl123/DBAddin) as well as other general Ribbon elements such as combobox, editbox, toggle button, split button, gallery (only with text) and checkbox.

To utilize placing dynamic command buttons in Excel worksheets you have to include Microsoft.Vbe.Interop.Forms.dll (in my case it is located in C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Vbe.Interop.Forms\11.0.0.0__71e9bce111e9429c\Microsoft.Vbe.Interop.Forms.dll) into this projects references.

After building the sample, simply activate the Addin by opening RibbonVB-AddIn64.xll or RibbonVB-AddIn.xll (32 bit) in bin/Debug or bin/Release and select the new RibbonVB Ribbon.