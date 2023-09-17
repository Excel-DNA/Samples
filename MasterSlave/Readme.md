The MasterSlave sample shows how an add-in host can control loading and unloading of other add-ins.

The Master add-in displays a ribbon with two buttons, to load and unload the Slave add-in.

In addition, there is TestController application (a console application) which loads and controls the Master add-in, and attempts to detect whether the Excel process exits with an error. This is to test for a reported error on closing Excel when an add-in has been unregistered.

Reloading of add-ins is not currently supported under .NET 6.

