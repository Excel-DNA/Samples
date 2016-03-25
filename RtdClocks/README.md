The RtdClocks sample shows a number of ways to implement RTD functions with Excel-DNA.

The RtdClocks solution has a number of projects - each project is a stand-alone add-in that shows one approach to implementing an RTD function.

## RtdClock-ExcelRtdServer
This project implements an RTD server using the Excel-DNA base class `ExcelRtdServer`. Any RTD server implemented in an Excel-DNA add-in should use the base class, rather than implementing the IRtdServer interface directly. The base class provides full access to all RTD features, and exposes a thread-safe and update notification that can be called at any time, at high frequency, from any thread.
Internally, Excel-DNA uses an ExcelRtdServer for all the other RTD-based features, including RxExcel / IObservable support.

## RtdClock-IExcelObservable
This project uses the higher-level abstraction of an IObservable / IObserver interface to implement the RTD function. In order to allow compatibility with .NET 2.0, the interfaces IExcelObservable / IExcelObserver are used by Excel-DNA, but the semantics is the same as the .NET 4.0 interfaces IObservable<object> / IObserver<object>.

## RtdClock-Rx
We show how to use the Reactive Extensions (Rx) library to define a simple timer. The Rx Observable is exposed as a UDF through the add-in using some helper utilities that are defined here.

## RtdClock-Rx-Registration
The Excel-DNA Registration extension library allows for custom registration extensions, including autoomatic registration of async and IObservable functions. This project shows how the registration of Rx / IObservable functions can be automatically done, exhibiting the cleanest minimal implementation of an RTD function.