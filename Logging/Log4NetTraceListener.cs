// This code is taken from the Cavity project (https://code.google.com/p/cavity/) by Alan Dean
// Retrieved on 2015-06-22 from:
// https://code.google.com/p/cavity/source/browse/trunk/src/Cavity%20Diagnostics%20log4net/Class%20Libraries/Diagnostics.Log4Net/Diagnostics/Log4NetTraceListener.cs
// Code License: MIT License (http://www.opensource.org/licenses/mit-license.php)

namespace Cavity.Diagnostics
{
    using System;
    using System.Diagnostics;
    using System.Diagnostics.CodeAnalysis;
    using System.Globalization;
    using System.Linq;
    using log4net;

    public class Log4NetTraceListener : TraceListener
    {
        public override void Fail(string message)
        {
            Fail(message, string.Empty);
        }

        public override void Fail(string message,
                                  string detailMessage)
        {
            var stack = new StackTrace();
            var frame = GetTracingStackFrame(stack);
            var log = LogManager.GetLogger(frame.GetMethod().DeclaringType);
            if (!log.IsWarnEnabled)
            {
                return;
            }

            using (ThreadContext.Stacks["signature"].Push(frame.GetMethod().Name))
            {
                message = string.IsNullOrEmpty(detailMessage)
                              ? message
                              : string.Format(CultureInfo.InvariantCulture, "{0}{1}{2}", message, Environment.NewLine, detailMessage);
                log.WarnFormat("[Fail] {0}", message);
            }
        }

        [SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId = "0", Justification = "Temporary")]
        public override void TraceData(TraceEventCache eventCache,
                                       string source,
                                       TraceEventType eventType,
                                       int id,
                                       object data)
        {
            var array = new[]
                            {
                                data
                            };
            TraceData(eventCache,
                      source,
                      eventType,
                      id,
                      array);
        }

        [SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId = "0", Justification = "Temporary")]
        public override void TraceData(TraceEventCache eventCache,
                                       string source,
                                       TraceEventType eventType,
                                       int id,
                                       params object[] data)
        {
            if (null == data ||
                0 == data.Length)
            {
                TraceEvent(eventCache, source, eventType, id);
                return;
            }

#if NET20
            foreach (var datum in data)
            {
                if (TraceException(eventType, datum))
                {
                    continue;
                }
            
                var array = new[]
                {
                    datum
                };
                TraceEvent(eventCache, 
                           source, 
                           eventType, 
                           id, 
                           "{0}", 
                           array);
            }
#else
            foreach (var datum in data.Where(x => !TraceException(eventType, x)))
            {
                var array = new[]
                                {
                                    datum
                                };
                TraceEvent(eventCache,
                           source,
                           eventType,
                           id,
                           "{0}",
                           array);
            }

#endif
        }

        [SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId = "0", Justification = "Temporary")]
        public override void TraceEvent(TraceEventCache eventCache,
                                        string source,
                                        TraceEventType eventType,
                                        int id)
        {
            var frame = GetTracingStackFrame(new StackTrace());
            var log = LogManager.GetLogger(frame.GetMethod().DeclaringType);
            switch (eventType)
            {
                case TraceEventType.Critical:
                    if (!log.IsFatalEnabled)
                    {
                        return;
                    }

                    using (ThreadContext.Stacks["signature"].Push(frame.GetMethod().Name))
                    {
                        log.Fatal(string.Empty);
                    }

                    break;

                case TraceEventType.Error:
                    if (!log.IsFatalEnabled)
                    {
                        return;
                    }

                    using (ThreadContext.Stacks["signature"].Push(frame.GetMethod().Name))
                    {
                        log.Error(string.Empty);
                    }

                    break;

                case TraceEventType.Information:
                    if (!log.IsInfoEnabled)
                    {
                        return;
                    }

                    using (ThreadContext.Stacks["signature"].Push(frame.GetMethod().Name))
                    {
                        log.Info(string.Empty);
                    }

                    break;

                case TraceEventType.Resume:
                case TraceEventType.Start:
                case TraceEventType.Stop:
                case TraceEventType.Suspend:
                case TraceEventType.Transfer:
                case TraceEventType.Verbose:
                    if (!log.IsDebugEnabled)
                    {
                        return;
                    }

                    using (ThreadContext.Stacks["signature"].Push(frame.GetMethod().Name))
                    {
                        log.Debug(string.Empty);
                    }

                    break;

                case TraceEventType.Warning:
                    if (!log.IsWarnEnabled)
                    {
                        return;
                    }

                    using (ThreadContext.Stacks["signature"].Push(frame.GetMethod().Name))
                    {
                        log.Warn(string.Empty);
                    }

                    break;
            }
        }

        [SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId = "0", Justification = "Argument cannot be null")]
        [SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId = "5", Justification = "Argument cannot be null")]
        public override void TraceEvent(TraceEventCache eventCache,
                                        string source,
                                        TraceEventType eventType,
                                        int id,
                                        string format,
            // ReSharper disable MethodOverloadWithOptionalParameter
                                        params object[] args)
        {
            // ReSharper restore MethodOverloadWithOptionalParameter
            if (0 == args.Length)
            {
                TraceEvent(eventCache, source, eventType, id);
            }

            var frame = GetTracingStackFrame(new StackTrace());
            var log = LogManager.GetLogger(frame.GetMethod().DeclaringType);
            using (ThreadContext.Stacks["signature"].Push(frame.GetMethod().Name))
            {
                switch (eventType)
                {
                    case TraceEventType.Critical:
                        if (!log.IsFatalEnabled)
                        {
                            return;
                        }

                        log.FatalFormat(CultureInfo.CurrentCulture, format, args);
                        break;

                    case TraceEventType.Error:
                        if (!log.IsFatalEnabled)
                        {
                            return;
                        }

                        log.ErrorFormat(CultureInfo.CurrentCulture, format, args);
                        break;

                    case TraceEventType.Information:
                        if (!log.IsInfoEnabled)
                        {
                            return;
                        }

                        log.InfoFormat(CultureInfo.CurrentCulture, format, args);
                        break;

                    case TraceEventType.Resume:
                    case TraceEventType.Start:
                    case TraceEventType.Stop:
                    case TraceEventType.Suspend:
                    case TraceEventType.Transfer:
                    case TraceEventType.Verbose:
                        if (!log.IsDebugEnabled)
                        {
                            return;
                        }

                        log.DebugFormat(CultureInfo.CurrentCulture, format, args);
                        break;

                    case TraceEventType.Warning:
                        if (!log.IsWarnEnabled)
                        {
                            return;
                        }

                        log.WarnFormat(CultureInfo.CurrentCulture, format, args);
                        break;
                }
            }
        }

        [SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId = "0", Justification = "Temporary")]
        public override void TraceEvent(TraceEventCache eventCache,
                                        string source,
                                        TraceEventType eventType,
                                        int id,
                                        string message)
        {
            var array = new[]
                            {
                                message
                            };
            TraceEvent(eventCache,
                       source,
                       eventType,
                       id,
                       "{0}",
                // ReSharper disable CoVariantArrayConversion
                       array);

            // ReSharper restore CoVariantArrayConversion
        }

        [SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId = "0", Justification = "Temporary")]
        public override void TraceTransfer(TraceEventCache eventCache,
                                           string source,
                                           int id,
                                           string message,
                                           Guid relatedActivityId)
        {
            var array = new object[]
                            {
                                message, relatedActivityId
                            };
            TraceEvent(eventCache,
                       source,
                       TraceEventType.Transfer,
                       id,
                       "{0}",
                       array);
        }

        public override void Write(object o)
        {
            WriteLine(o, string.Empty);
        }

        public override void Write(object o,
                                   string category)
        {
            WriteLine(o, category);
        }

        public override void Write(string message)
        {
            WriteLine(message, string.Empty);
        }

        public override void Write(string message,
                                   string category)
        {
            WriteLine((object)message, category);
        }

        public override void WriteLine(object o)
        {
            WriteLine(o, string.Empty);
        }

        [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "log4net.Util.ThreadContextStack.Push(System.String)", Justification = "Temporary")]
        public override void WriteLine(object o,
                                       string category)
        {
            var stack = new StackTrace();
            var frame = GetTracingStackFrame(stack);
            var log = LogManager.GetLogger(frame.GetMethod().DeclaringType);
            if (!log.IsInfoEnabled)
            {
                return;
            }

            using (ThreadContext.Stacks["signature"].Push(frame.GetMethod().Name))
            {
                if (string.IsNullOrEmpty(category))
                {
                    log.Debug(o);
                }
                else
                {
                    log.DebugFormat("[{0}] {1}", category, o);
                }
            }
        }

        public override void WriteLine(string message,
                                       string category)
        {
            WriteLine((object)message, category);
        }

        public override void WriteLine(string message)
        {
            WriteLine(message, string.Empty);
        }

        private static StackFrame GetTracingStackFrame(StackTrace stack)
        {
            for (var i = 0; i < stack.FrameCount; i++)
            {
                var frame = stack.GetFrame(i);
                var method = frame.GetMethod();
                if (null == method)
                {
                    continue;
                }

                // ReSharper disable PossibleNullReferenceException
                if ("System.Diagnostics" == method.DeclaringType.Namespace)
                {
                    // ReSharper restore PossibleNullReferenceException
                    continue;
                }

                if ("System.Threading" == method.DeclaringType.Namespace)
                {
                    continue;
                }

                if ("Log4NetTraceListener" == method.DeclaringType.Name)
                {
                    continue;
                }

                return stack.GetFrame(i);
            }

            return null;
        }

        private static bool TraceException(TraceEventType eventType,
                                           object datum)
        {
            if (TraceEventType.Critical != eventType &&
                TraceEventType.Error != eventType)
            {
                return false;
            }

            var exception = datum as Exception;
            if (null == exception)
            {
                return false;
            }

            var frame = GetTracingStackFrame(new StackTrace());
            var log = LogManager.GetLogger(frame.GetMethod().DeclaringType);
            switch (eventType)
            {
                case TraceEventType.Critical:
                    if (log.IsFatalEnabled)
                    {
                        using (ThreadContext.Stacks["signature"].Push(frame.GetMethod().Name))
                        {
                            log.Fatal(exception.Message, exception);
                        }
                    }

                    break;

                case TraceEventType.Error:
                    if (log.IsErrorEnabled)
                    {
                        using (ThreadContext.Stacks["signature"].Push(frame.GetMethod().Name))
                        {
                            log.Error(exception.Message, exception);
                        }
                    }

                    break;
            }

            return true;
        }
    }
}