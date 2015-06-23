Excel-DNA Logging Samples
=========================

The Logging samples show different ways of configuring the Excel-DNA diagnostic logging.
To examine each option, uncomment the corresponding part in the Logging-AddIn.xll.config file.

### log4net
A Log4Net TraceListener is defined (in this example project), that forwards trace events to Log4Net.

### NLog
The NLog library includes an NLogTraceListener that can be configured to forward trace events to NLog.